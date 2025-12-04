
<#
.SYNOPSIS
    Ultimate Microsoft 365 Management GUI
.DESCRIPTION
    A comprehensive Windows Forms GUI for managing Microsoft 365 tenant connections and user offboarding.
    Features include: Connect to multiple services, manage Graph API scopes, and perform offboarding tasks.
.NOTES
    Author: Seth G
    Version: 1.8 - Enhanced Error Handling (Updated 12-03-2024)
    Requires: Windows PowerShell 5.1 or PowerShell 7+
    
    SAVE THIS FILE AS: UltimateM365Management-v1.8.ps1
    
    CHANGES IN v1.8:
    - Improved Exchange Online connection with timeout handling
    - Better module import error handling
    - Enhanced connection verification with Get-ConnectionInformation
    - Added retry logic and connection cleanup
    - More detailed error messages for troubleshooting
    - Improved logging for connection issues
#>

#Requires -Version 5.1

# Import required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# Initialize logging
$Global:LogPath = "C:\Temp\UltimatePS"
if (-not (Test-Path $Global:LogPath)) {
    New-Item -Path $Global:LogPath -ItemType Directory -Force | Out-Null
}
$Global:LogFile = Join-Path $Global:LogPath "UltimatePS_$(Get-Date -Format 'yyyyMMdd').log"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Add-Content -Path $Global:LogFile -Value $logMessage
}

function Show-ErrorMessage {
    param(
        [string]$Title,
        [string]$Message
    )
    Write-Log -Message "$Title - $Message" -Level ERROR
    [System.Windows.Forms.MessageBox]::Show($Message, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
}

function Show-InfoMessage {
    param(
        [string]$Title,
        [string]$Message
    )
    Write-Log -Message "$Title - $Message" -Level INFO
    [System.Windows.Forms.MessageBox]::Show($Message, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}

# Global connection status
$Global:ConnectionStatus = @{
    Graph = @{ Connected = $false; Tenant = "Not Connected"; LastCheck = $null }
    Exchange = @{ Connected = $false; Tenant = "Not Connected"; LastCheck = $null }
    MSOnline = @{ Connected = $false; Tenant = "Not Connected"; LastCheck = $null }
    AzureAD = @{ Connected = $false; Tenant = "Not Connected"; LastCheck = $null }
    SharePoint = @{ Connected = $false; Tenant = "Not Connected"; LastCheck = $null }
    Teams = @{ Connected = $false; Tenant = "Not Connected"; LastCheck = $null }
}

# Default Graph API Scopes
$Global:GraphScopes = @(
    "User.ReadWrite.All",
    "Mail.ReadWrite",
    "MailboxSettings.ReadWrite",
    "Directory.ReadWrite.All"
)

# License SKU ID to Friendly Name Mapping
$Global:LicenseNameMap = @{
    "f245ecc8-75af-4f8e-b61f-27d8114de5f3" = "Microsoft 365 Business Standard"
    "4b9405b0-7788-4568-add1-99614e613b69" = "Exchange Online Plan 1"
    "3b555118-da6a-4418-894f-7df1e2096870" = "Microsoft 365 Business Basic"
    "078d2b04-f1bd-4111-bbd4-b4b1b354cef4" = "Azure Active Directory Premium P1"
    "6470687e-a428-4b7a-bef2-8a291ad125d4" = "Microsoft 365 Business Premium"
    "05e9a617-0261-4cee-bb44-138d3ef2d965" = "Microsoft 365 Apps for Business"
    "1f2f344a-700d-42c9-9427-5cea1d5d7ba7" = "Microsoft 365 Business Enterprise"
    "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46" = "Enterprise Mobility + Security E3"
    "b05e124f-c7cc-45a0-a6aa-8cf78c946968" = "Enterprise Mobility + Security E5"
}

function Test-AndInstallModule {
    param([string]$ModuleName)
    
    Write-Log "Checking module: $ModuleName" -Level DEBUG
    
    if (Get-Module -Name $ModuleName -ErrorAction SilentlyContinue) {
        Write-Log "Module '$ModuleName' is already loaded"
        return $true
    }
    
    $availableModule = Get-Module -ListAvailable -Name $ModuleName -ErrorAction SilentlyContinue
    if ($availableModule) {
        Write-Log "Module '$ModuleName' found, attempting to import..."
        try {
            Import-Module -Name $ModuleName -Force -ErrorAction Stop
            Write-Log "Successfully imported module '$ModuleName'"
            return $true
        }
        catch {
            Write-Log "Failed to import module '$ModuleName': $_" -Level "WARNING"
        }
    }
    
    Write-Log "Module '$ModuleName' not available, prompting user for installation"
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Module '$ModuleName' is not installed. Would you like to install it now?",
        "Module Required",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq 'Yes') {
        try {
            Write-Log "Installing module: $ModuleName"
            Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
            Import-Module -Name $ModuleName -Force -ErrorAction Stop
            Write-Log "Module '$ModuleName' installed successfully"
            Show-InfoMessage -Title "Success" -Message "Module '$ModuleName' installed successfully."
            return $true
        }
        catch {
            Write-Log "Failed to install '$ModuleName': $($_.Exception.Message)" -Level "ERROR"
            Show-ErrorMessage -Title "Installation Failed" -Message "Failed to install '$ModuleName':`n`n$($_.Exception.Message)"
            return $false
        }
    }
    return $false
}

function Connect-ToGraph {
    try {
        Write-Log "=== Starting Microsoft Graph connection ===" -Level DEBUG
        Import-Module -Name Microsoft.Graph.Authentication -Force -ErrorAction Stop
        Import-Module -Name Microsoft.Graph.Users -Force -ErrorAction Stop
        
        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($context) {
            $tenantId = $context.TenantId
            Write-Log "Already connected to Microsoft Graph - Tenant: $tenantId"
            $Global:ConnectionStatus.Graph.Connected = $true
            $Global:ConnectionStatus.Graph.Tenant = $tenantId
            $Global:ConnectionStatus.Graph.LastCheck = Get-Date
            Update-ConnectionStatus
            Show-InfoMessage -Title "Already Connected" -Message "Already connected to Microsoft Graph`nTenant: $tenantId"
            return $true
        }

        Write-Log "Connecting to Microsoft Graph..."
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch { }
        Connect-MgGraph -Scopes $Global:GraphScopes -ErrorAction Stop
        
        $context = Get-MgContext -ErrorAction Stop
        if ($context) {
            $tenantId = $context.TenantId
            $Global:ConnectionStatus.Graph.Connected = $true
            $Global:ConnectionStatus.Graph.Tenant = $tenantId
            $Global:ConnectionStatus.Graph.LastCheck = Get-Date
            Write-Log "Successfully connected to Microsoft Graph"
            Update-ConnectionStatus
            Show-InfoMessage -Title "Success" -Message "Connected to Microsoft Graph`nTenant: $tenantId"
            return $true
        }
    }
    catch {
        Write-Log "Graph connection failed: $($_.Exception.Message)" -Level ERROR
        $Global:ConnectionStatus.Graph.Connected = $false
        $Global:ConnectionStatus.Graph.Tenant = "Connection Failed"
        Update-ConnectionStatus
        Show-ErrorMessage -Title "Connection Failed" -Message "Failed to connect to Microsoft Graph:`n`n$($_.Exception.Message)"
        return $false
    }
}

function Connect-ToExchange {
    try {
        Write-Log "=== Starting Exchange Online connection ===" -Level DEBUG
        
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            throw "ExchangeOnlineManagement module not installed. Please install it: Install-Module -Name ExchangeOnlineManagement"
        }
        
        Import-Module -Name ExchangeOnlineManagement -Force -ErrorAction Stop
        Write-Log "ExchangeOnlineManagement module loaded"
        
        Write-Log "Checking for existing connection..."
        try {
            $existingConnection = Get-ConnectionInformation -ErrorAction Stop 2>$null
            if ($existingConnection -and $existingConnection.State -eq 'Connected') {
                Write-Log "Already connected - Organization: $($existingConnection.Organization)"
                $Global:ConnectionStatus.Exchange.Connected = $true
                $Global:ConnectionStatus.Exchange.Tenant = $existingConnection.Organization
                $Global:ConnectionStatus.Exchange.LastCheck = Get-Date
                Update-ConnectionStatus
                Show-InfoMessage -Title "Already Connected" -Message "Already connected to Exchange Online`nOrg: $($existingConnection.Organization)"
                return $true
            }
        } catch {
            Write-Log "No existing connection found" -Level DEBUG
        }

        Write-Log "Connecting to Exchange Online..."
        
        # Clean disconnect first
        try { 
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue 2>$null
        } catch { }
        
        # Connect with interactive login (like Graph API)
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        
        # Verify connection
        Start-Sleep -Seconds 1
        
        $connectionInfo = Get-ConnectionInformation -ErrorAction Stop 2>$null
        if ($connectionInfo -and $connectionInfo.State -eq 'Connected') {
            $orgName = $connectionInfo.Organization
            Write-Log "Connection verified - Organization: $orgName"
            $Global:ConnectionStatus.Exchange.Connected = $true
            $Global:ConnectionStatus.Exchange.Tenant = $orgName
            $Global:ConnectionStatus.Exchange.LastCheck = Get-Date
            Update-ConnectionStatus
            Show-InfoMessage -Title "Success" -Message "Connected to Exchange Online`nOrganization: $orgName"
            return $true
        } else {
            throw "Connection verification failed"
        }
    }
    catch {
        $errorDetails = $_.Exception.Message
        Write-Log "Exchange connection failed: $errorDetails" -Level ERROR
        
        $userMessage = "Failed to connect to Exchange Online:`n`n$errorDetails"
        
        if ($errorDetails -like "*authentication*" -or $errorDetails -like "*credential*") {
            $userMessage += "`n`n💡 Check credentials and MFA"
        }
        elseif ($errorDetails -like "*user*cancel*") {
            $userMessage += "`n`n💡 Authentication was cancelled"
        }
        
        $Global:ConnectionStatus.Exchange.Connected = $false
        $Global:ConnectionStatus.Exchange.Tenant = "Connection Failed"
        Update-ConnectionStatus
        Show-ErrorMessage -Title "Connection Failed" -Message $userMessage
        return $false
    }
}
function Connect-ToMSOnline {
    try {
        Import-Module -Name MSOnline -Force -ErrorAction Stop
        Connect-MsolService -ErrorAction Stop
        $tenant = Get-MsolCompanyInformation -ErrorAction Stop
        if ($tenant) {
            $Global:ConnectionStatus.MSOnline.Connected = $true
            $Global:ConnectionStatus.MSOnline.Tenant = $tenant.DisplayName
            $Global:ConnectionStatus.MSOnline.LastCheck = Get-Date
            Write-Log "Connected to MSOnline - Tenant: $($tenant.DisplayName)"
            Update-ConnectionStatus
            Show-InfoMessage -Title "Success" -Message "Connected to MSOnline`nTenant: $($tenant.DisplayName)"
        }
    }
    catch {
        Write-Log "MSOnline connection failed: $($_.Exception.Message)" -Level ERROR
        Show-ErrorMessage -Title "Connection Failed" -Message $_.Exception.Message
    }
}

function Connect-ToAzureAD {
    try {
        Import-Module -Name AzureAD -Force -ErrorAction Stop
        $connection = Connect-AzureAD -ErrorAction Stop
        if ($connection) {
            $tenant = Get-AzureADTenantDetail
            $Global:ConnectionStatus.AzureAD.Connected = $true
            $Global:ConnectionStatus.AzureAD.Tenant = $tenant.DisplayName
            $Global:ConnectionStatus.AzureAD.LastCheck = Get-Date
            Write-Log "Connected to Azure AD - Tenant: $($tenant.DisplayName)"
            Update-ConnectionStatus
            Show-InfoMessage -Title "Success" -Message "Connected to Azure AD`nTenant: $($tenant.DisplayName)"
        }
    }
    catch {
        Write-Log "Azure AD connection failed: $($_.Exception.Message)" -Level ERROR
        Show-ErrorMessage -Title "Connection Failed" -Message $_.Exception.Message
    }
}

function Connect-ToSharePoint {
    try {
        Import-Module -Name Microsoft.Online.SharePoint.PowerShell -Force -ErrorAction Stop
        $adminUrl = [Microsoft.VisualBasic.Interaction]::InputBox("Enter SharePoint Admin URL", "SharePoint Admin URL", "https://")
        if ([string]::IsNullOrWhiteSpace($adminUrl)) { return }
        Connect-SPOService -Url $adminUrl -ErrorAction Stop
        $Global:ConnectionStatus.SharePoint.Connected = $true
        $Global:ConnectionStatus.SharePoint.Tenant = $adminUrl
        $Global:ConnectionStatus.SharePoint.LastCheck = Get-Date
        Write-Log "Connected to SharePoint Online"
        Update-ConnectionStatus
        Show-InfoMessage -Title "Success" -Message "Connected to SharePoint Online"
    }
    catch {
        Write-Log "SharePoint connection failed: $($_.Exception.Message)" -Level ERROR
        Show-ErrorMessage -Title "Connection Failed" -Message $_.Exception.Message
    }
}

function Connect-ToTeams {
    try {
        Import-Module -Name MicrosoftTeams -Force -ErrorAction Stop
        try { Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue } catch { }
        Connect-MicrosoftTeams -ErrorAction Stop
        $tenant = Get-CsTenant -ErrorAction Stop
        if ($tenant) {
            $Global:ConnectionStatus.Teams.Connected = $true
            $Global:ConnectionStatus.Teams.Tenant = $tenant.DisplayName
            $Global:ConnectionStatus.Teams.LastCheck = Get-Date
            Write-Log "Connected to Teams - Tenant: $($tenant.DisplayName)"
            Update-ConnectionStatus
            Show-InfoMessage -Title "Success" -Message "Connected to Microsoft Teams`nTenant: $($tenant.DisplayName)"
            return $true
        }
    }
    catch {
        Write-Log "Teams connection failed: $($_.Exception.Message)" -Level ERROR
        $Global:ConnectionStatus.Teams.Connected = $false
        $Global:ConnectionStatus.Teams.Tenant = "Connection Failed"
        Update-ConnectionStatus
        Show-ErrorMessage -Title "Connection Failed" -Message $_.Exception.Message
        return $false
    }
}

# Disconnect functions
function Disconnect-FromGraph {
    try {
        if ($Global:ConnectionStatus.Graph.Connected) {
            Disconnect-MgGraph -ErrorAction Stop | Out-Null
            $Global:ConnectionStatus.Graph.Connected = $false
            $Global:ConnectionStatus.Graph.Tenant = "Not Connected"
            Write-Log "Disconnected from Graph"
            Update-ConnectionStatus
            Show-InfoMessage -Title "Success" -Message "Disconnected from Microsoft Graph"
        }
    }
    catch { Show-ErrorMessage -Title "Disconnect Failed" -Message $_.Exception.Message }
}

function Disconnect-FromExchange {
    try {
        if ($Global:ConnectionStatus.Exchange.Connected) {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
            $Global:ConnectionStatus.Exchange.Connected = $false
            $Global:ConnectionStatus.Exchange.Tenant = "Not Connected"
            Write-Log "Disconnected from Exchange"
            Update-ConnectionStatus
            Show-InfoMessage -Title "Success" -Message "Disconnected from Exchange Online"
        }
    }
    catch { Show-ErrorMessage -Title "Disconnect Failed" -Message $_.Exception.Message }
}

function Disconnect-FromMSOnline {
    try {
        if ($Global:ConnectionStatus.MSOnline.Connected) {
            $Global:ConnectionStatus.MSOnline.Connected = $false
            $Global:ConnectionStatus.MSOnline.Tenant = "Not Connected"
            Update-ConnectionStatus
            Show-InfoMessage -Title "Info" -Message "MSOnline status cleared"
        }
    }
    catch { Show-ErrorMessage -Title "Disconnect Failed" -Message $_.Exception.Message }
}

function Disconnect-FromAzureAD {
    try {
        if ($Global:ConnectionStatus.AzureAD.Connected) {
            Disconnect-AzureAD -ErrorAction Stop
            $Global:ConnectionStatus.AzureAD.Connected = $false
            $Global:ConnectionStatus.AzureAD.Tenant = "Not Connected"
            Update-ConnectionStatus
            Show-InfoMessage -Title "Success" -Message "Disconnected from Azure AD"
        }
    }
    catch { Show-ErrorMessage -Title "Disconnect Failed" -Message $_.Exception.Message }
}

function Disconnect-FromSharePoint {
    try {
        if ($Global:ConnectionStatus.SharePoint.Connected) {
            Disconnect-SPOService -ErrorAction Stop
            $Global:ConnectionStatus.SharePoint.Connected = $false
            $Global:ConnectionStatus.SharePoint.Tenant = "Not Connected"
            Update-ConnectionStatus
            Show-InfoMessage -Title "Success" -Message "Disconnected from SharePoint"
        }
    }
    catch { Show-ErrorMessage -Title "Disconnect Failed" -Message $_.Exception.Message }
}

function Disconnect-FromTeams {
    try {
        if ($Global:ConnectionStatus.Teams.Connected) {
            Disconnect-MicrosoftTeams -ErrorAction Stop
            $Global:ConnectionStatus.Teams.Connected = $false
            $Global:ConnectionStatus.Teams.Tenant = "Not Connected"
            Update-ConnectionStatus
            Show-InfoMessage -Title "Success" -Message "Disconnected from Teams"
        }
    }
    catch { Show-ErrorMessage -Title "Disconnect Failed" -Message $_.Exception.Message }
}

# Test connection functions
function Test-GraphConnection {
    try {
        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($context) {
            $Global:ConnectionStatus.Graph.Connected = $true
            $Global:ConnectionStatus.Graph.Tenant = $context.TenantId
            $Global:ConnectionStatus.Graph.LastCheck = Get-Date
            return $true
        }
        $Global:ConnectionStatus.Graph.Connected = $false
        return $false
    }
    catch { $Global:ConnectionStatus.Graph.Connected = $false; return $false }
}

function Test-ExchangeConnection {
    try {
        $connectionInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue 2>$null
        if ($connectionInfo -and $connectionInfo.State -eq 'Connected') {
            $Global:ConnectionStatus.Exchange.Connected = $true
            $Global:ConnectionStatus.Exchange.Tenant = $connectionInfo.Organization
            $Global:ConnectionStatus.Exchange.LastCheck = Get-Date
            return $true
        }
        $org = Get-OrganizationConfig -ErrorAction SilentlyContinue 2>$null
        if ($org) {
            $Global:ConnectionStatus.Exchange.Connected = $true
            $Global:ConnectionStatus.Exchange.Tenant = $org.Name
            $Global:ConnectionStatus.Exchange.LastCheck = Get-Date
            return $true
        }
        $Global:ConnectionStatus.Exchange.Connected = $false
        return $false
    }
    catch { $Global:ConnectionStatus.Exchange.Connected = $false; return $false }
}

function Test-MSOnlineConnection {
    try {
        $tenant = Get-MsolCompanyInformation -ErrorAction SilentlyContinue
        if ($tenant) {
            $Global:ConnectionStatus.MSOnline.Connected = $true
            $Global:ConnectionStatus.MSOnline.Tenant = $tenant.DisplayName
            $Global:ConnectionStatus.MSOnline.LastCheck = Get-Date
            return $true
        }
        $Global:ConnectionStatus.MSOnline.Connected = $false
        return $false
    }
    catch { $Global:ConnectionStatus.MSOnline.Connected = $false; return $false }
}

function Test-AzureADConnection {
    try {
        $tenant = Get-AzureADTenantDetail -ErrorAction SilentlyContinue
        if ($tenant) {
            $Global:ConnectionStatus.AzureAD.Connected = $true
            $Global:ConnectionStatus.AzureAD.Tenant = $tenant.DisplayName
            $Global:ConnectionStatus.AzureAD.LastCheck = Get-Date
            return $true
        }
        $Global:ConnectionStatus.AzureAD.Connected = $false
        return $false
    }
    catch { $Global:ConnectionStatus.AzureAD.Connected = $false; return $false }
}

function Test-SharePointConnection {
    try {
        if ($Global:ConnectionStatus.SharePoint.Connected) {
            $Global:ConnectionStatus.SharePoint.LastCheck = Get-Date
            return $true
        }
        return $false
    }
    catch { $Global:ConnectionStatus.SharePoint.Connected = $false; return $false }
}

function Test-TeamsConnection {
    try {
        $tenant = Get-CsTenant -ErrorAction SilentlyContinue
        if ($tenant) {
            $Global:ConnectionStatus.Teams.Connected = $true
            $Global:ConnectionStatus.Teams.Tenant = $tenant.DisplayName
            $Global:ConnectionStatus.Teams.LastCheck = Get-Date
            return $true
        }
        $Global:ConnectionStatus.Teams.Connected = $false
        return $false
    }
    catch { $Global:ConnectionStatus.Teams.Connected = $false; return $false }
}

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Ultimate Microsoft 365 Management Tool v1.8"
$form.Size = New-Object System.Drawing.Size(900, 700)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$tabControl.Size = New-Object System.Drawing.Size(860, 640)
$form.Controls.Add($tabControl)

# TAB 1: Connect to Tenant
$tabConnect = New-Object System.Windows.Forms.TabPage
$tabConnect.Text = "Connect to Tenant"
$tabControl.Controls.Add($tabConnect)

$yPosition = 20
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Location = New-Object System.Drawing.Point(20, $yPosition)
$lblTitle.Size = New-Object System.Drawing.Size(800, 30)
$lblTitle.Text = "Microsoft 365 Service Connections"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$tabConnect.Controls.Add($lblTitle)
$yPosition += 50

$services = @(
    @{Name = "Microsoft Graph"; ConnectFunc = "Connect-ToGraph"; DisconnectFunc = "Disconnect-FromGraph"; StatusKey = "Graph"},
    @{Name = "Exchange Online"; ConnectFunc = "Connect-ToExchange"; DisconnectFunc = "Disconnect-FromExchange"; StatusKey = "Exchange"},
    @{Name = "MSOnline"; ConnectFunc = "Connect-ToMSOnline"; DisconnectFunc = "Disconnect-FromMSOnline"; StatusKey = "MSOnline"},
    @{Name = "Azure AD"; ConnectFunc = "Connect-ToAzureAD"; DisconnectFunc = "Disconnect-FromAzureAD"; StatusKey = "AzureAD"},
    @{Name = "SharePoint Online"; ConnectFunc = "Connect-ToSharePoint"; DisconnectFunc = "Disconnect-FromSharePoint"; StatusKey = "SharePoint"},
    @{Name = "Microsoft Teams"; ConnectFunc = "Connect-ToTeams"; DisconnectFunc = "Disconnect-FromTeams"; StatusKey = "Teams"}
)

$Global:StatusLabels = @{}

foreach ($service in $services) {
    $lblService = New-Object System.Windows.Forms.Label
    $lblService.Location = New-Object System.Drawing.Point(20, $yPosition)
    $lblService.Size = New-Object System.Drawing.Size(150, 25)
    $lblService.Text = "$($service.Name):"
    $lblService.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $tabConnect.Controls.Add($lblService)
    
    $btnConnect = New-Object System.Windows.Forms.Button
    $btnConnect.Location = New-Object System.Drawing.Point(180, $yPosition)
    $btnConnect.Size = New-Object System.Drawing.Size(100, 25)
    $btnConnect.Text = "Connect"
    $btnConnect.Add_Click({param($btn, $e); & $btn.Tag}.GetNewClosure())
    $btnConnect.Tag = $service.ConnectFunc
    $tabConnect.Controls.Add($btnConnect)
    
    $btnDisconnect = New-Object System.Windows.Forms.Button
    $btnDisconnect.Location = New-Object System.Drawing.Point(290, $yPosition)
    $btnDisconnect.Size = New-Object System.Drawing.Size(100, 25)
    $btnDisconnect.Text = "Disconnect"
    $btnDisconnect.Add_Click({param($btn, $e); & $btn.Tag}.GetNewClosure())
    $btnDisconnect.Tag = $service.DisconnectFunc
    $tabConnect.Controls.Add($btnDisconnect)
    
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(400, $yPosition)
    $lblStatus.Size = New-Object System.Drawing.Size(430, 25)
    $lblStatus.Text = "Status: Not Connected | Tenant: N/A"
    $lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblStatus.ForeColor = [System.Drawing.Color]::Red
    $tabConnect.Controls.Add($lblStatus)
    
    $Global:StatusLabels[$service.StatusKey] = $lblStatus
    $yPosition += 40
}

$statusTimer = New-Object System.Windows.Forms.Timer
$statusTimer.Interval = 30000
$statusTimer.Add_Tick({
    Test-GraphConnection | Out-Null
    Test-ExchangeConnection | Out-Null
    Test-MSOnlineConnection | Out-Null
    Test-AzureADConnection | Out-Null
    Test-SharePointConnection | Out-Null
    Test-TeamsConnection | Out-Null
    Update-ConnectionStatus
})
$statusTimer.Start()

$btnRefreshStatus = New-Object System.Windows.Forms.Button
$btnRefreshStatus.Location = New-Object System.Drawing.Point(20, ($yPosition + 20))
$btnRefreshStatus.Size = New-Object System.Drawing.Size(150, 30)
$btnRefreshStatus.Text = "Refresh Status Now"
$btnRefreshStatus.Add_Click({
    Test-GraphConnection | Out-Null
    Test-ExchangeConnection | Out-Null
    Test-MSOnlineConnection | Out-Null
    Test-AzureADConnection | Out-Null
    Test-SharePointConnection | Out-Null
    Test-TeamsConnection | Out-Null
    Update-ConnectionStatus
    Show-InfoMessage -Title “Status Updated” -Message “Connection status refreshed successfully.”
})
$tabConnect.Controls.Add($btnRefreshStatus)

$Global:lblLastCheck = New-Object System.Windows.Forms.Label
$Global:lblLastCheck.Location = New-Object System.Drawing.Point(180, ($yPosition + 25))
$Global:lblLastCheck.Size = New-Object System.Drawing.Size(400, 20)
$Global:lblLastCheck.Text = “Auto-refresh every 30 seconds”
$Global:lblLastCheck.Font = New-Object System.Drawing.Font(“Segoe UI”, 8, [System.Drawing.FontStyle]::Italic)
$Global:lblLastCheck.ForeColor = [System.Drawing.Color]::Gray
$tabConnect.Controls.Add($Global:lblLastCheck)

# TAB 2: Graph API Scope

$tabGraphScope = New-Object System.Windows.Forms.TabPage
$tabGraphScope.Text = “Graph API Scope”
$tabControl.Controls.Add($tabGraphScope)

$yPosScope = 20
$lblScopeTitle = New-Object System.Windows.Forms.Label
$lblScopeTitle.Location = New-Object System.Drawing.Point(20, $yPosScope)
$lblScopeTitle.Size = New-Object System.Drawing.Size(800, 30)
$lblScopeTitle.Text = “Configure Microsoft Graph API Permissions”
$lblScopeTitle.Font = New-Object System.Drawing.Font(“Segoe UI”, 14, [System.Drawing.FontStyle]::Bold)
$tabGraphScope.Controls.Add($lblScopeTitle)
$yPosScope += 50

$lblScopeInstructions = New-Object System.Windows.Forms.Label
$lblScopeInstructions.Location = New-Object System.Drawing.Point(20, $yPosScope)
$lblScopeInstructions.Size = New-Object System.Drawing.Size(800, 40)
$lblScopeInstructions.Text = “Select the Graph API permissions you need. After selecting, click ‘Apply & Reconnect’ to reconnect with new permissions.”
$lblScopeInstructions.Font = New-Object System.Drawing.Font(“Segoe UI”, 9)
$tabGraphScope.Controls.Add($lblScopeInstructions)
$yPosScope += 60

$availableScopes = @(
“User.Read.All”, “User.ReadWrite.All”, “Directory.Read.All”, “Directory.ReadWrite.All”,
“Group.Read.All”, “Group.ReadWrite.All”, “Mail.Read”, “Mail.ReadWrite”, “Mail.Send”,
“MailboxSettings.Read”, “MailboxSettings.ReadWrite”, “Calendars.Read”, “Calendars.ReadWrite”,
“Contacts.Read”, “Contacts.ReadWrite”, “Files.Read.All”, “Files.ReadWrite.All”,
“Sites.Read.All”, “Sites.ReadWrite.All”, “TeamSettings.Read.All”, “TeamSettings.ReadWrite.All”,
“Channel.ReadBasic.All”, “ChannelSettings.ReadWrite.All”, “Device.Read.All”, “Device.ReadWrite.All”
)

$checkedListScopes = New-Object System.Windows.Forms.CheckedListBox
$checkedListScopes.Location = New-Object System.Drawing.Point(20, $yPosScope)
$checkedListScopes.Size = New-Object System.Drawing.Size(600, 350)
$checkedListScopes.CheckOnClick = $true
$checkedListScopes.Font = New-Object System.Drawing.Font(“Segoe UI”, 9)
$tabGraphScope.Controls.Add($checkedListScopes)

foreach ($scope in $availableScopes | Sort-Object) {
$index = $checkedListScopes.Items.Add($scope)
if ($Global:GraphScopes -contains $scope) {
$checkedListScopes.SetItemChecked($index, $true)
}
}
$yPosScope += 370

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Location = New-Object System.Drawing.Point(20, $yPosScope)
$btnSelectAll.Size = New-Object System.Drawing.Size(100, 30)
$btnSelectAll.Text = “Select All”
$btnSelectAll.Add_Click({
for ($i = 0; $i -lt $checkedListScopes.Items.Count; $i++) {
$checkedListScopes.SetItemChecked($i, $true)
}
})
$tabGraphScope.Controls.Add($btnSelectAll)

$btnClearAll = New-Object System.Windows.Forms.Button
$btnClearAll.Location = New-Object System.Drawing.Point(130, $yPosScope)
$btnClearAll.Size = New-Object System.Drawing.Size(100, 30)
$btnClearAll.Text = “Clear All”
$btnClearAll.Add_Click({
for ($i = 0; $i -lt $checkedListScopes.Items.Count; $i++) {
$checkedListScopes.SetItemChecked($i, $false)
}
})
$tabGraphScope.Controls.Add($btnClearAll)

$btnApplyScopes = New-Object System.Windows.Forms.Button
$btnApplyScopes.Location = New-Object System.Drawing.Point(250, $yPosScope)
$btnApplyScopes.Size = New-Object System.Drawing.Size(180, 30)
$btnApplyScopes.Text = “Apply & Reconnect to Graph”
$btnApplyScopes.Font = New-Object System.Drawing.Font(“Segoe UI”, 9, [System.Drawing.FontStyle]::Bold)
$btnApplyScopes.BackColor = [System.Drawing.Color]::LightGreen
$btnApplyScopes.Add_Click({
$selectedScopes = @()
foreach ($item in $checkedListScopes.CheckedItems) { $selectedScopes += $item }


if ($selectedScopes.Count -eq 0) {
    Show-ErrorMessage -Title "No Scopes Selected" -Message "Please select at least one permission scope."
    return
}

$Global:GraphScopes = $selectedScopes
Write-Log "Updated Graph API scopes: $($Global:GraphScopes -join ', ')"

if ($Global:ConnectionStatus.Graph.Connected) { Disconnect-FromGraph }
Connect-ToGraph


})
$tabGraphScope.Controls.Add($btnApplyScopes)

$Global:lblCurrentScopes = New-Object System.Windows.Forms.Label
$Global:lblCurrentScopes.Location = New-Object System.Drawing.Point(20, ($yPosScope + 40))
$Global:lblCurrentScopes.Size = New-Object System.Drawing.Size(600, 60)
$Global:lblCurrentScopes.Text = “Current Scopes: $($Global:GraphScopes -join ’, ’)”
$Global:lblCurrentScopes.Font = New-Object System.Drawing.Font(“Segoe UI”, 8)
$Global:lblCurrentScopes.ForeColor = [System.Drawing.Color]::DarkBlue
$tabGraphScope.Controls.Add($Global:lblCurrentScopes)

# TAB 3: Offboarding

$tabOffboarding = New-Object System.Windows.Forms.TabPage
$tabOffboarding.Text = “Offboarding”
$tabOffboarding.AutoScroll = $true
$tabControl.Controls.Add($tabOffboarding)

$yPosOff = 20
$lblOffTitle = New-Object System.Windows.Forms.Label
$lblOffTitle.Location = New-Object System.Drawing.Point(20, $yPosOff)
$lblOffTitle.Size = New-Object System.Drawing.Size(800, 30)
$lblOffTitle.Text = “User Offboarding Tools”
$lblOffTitle.Font = New-Object System.Drawing.Font(“Segoe UI”, 14, [System.Drawing.FontStyle]::Bold)
$tabOffboarding.Controls.Add($lblOffTitle)
$yPosOff += 50

$lblUserSearch = New-Object System.Windows.Forms.Label
$lblUserSearch.Location = New-Object System.Drawing.Point(20, $yPosOff)
$lblUserSearch.Size = New-Object System.Drawing.Size(150, 25)
$lblUserSearch.Text = “Search User (UPN):”
$lblUserSearch.Font = New-Object System.Drawing.Font(“Segoe UI”, 10, [System.Drawing.FontStyle]::Bold)
$tabOffboarding.Controls.Add($lblUserSearch)

$txtUserSearch = New-Object System.Windows.Forms.TextBox
$txtUserSearch.Location = New-Object System.Drawing.Point(180, $yPosOff)
$txtUserSearch.Size = New-Object System.Drawing.Size(400, 25)
$txtUserSearch.Font = New-Object System.Drawing.Font(“Segoe UI”, 10)
$tabOffboarding.Controls.Add($txtUserSearch)

$btnSearchUser = New-Object System.Windows.Forms.Button
$btnSearchUser.Location = New-Object System.Drawing.Point(590, $yPosOff)
$btnSearchUser.Size = New-Object System.Drawing.Size(100, 25)
$btnSearchUser.Text = “Search”
$btnSearchUser.Add_Click({
if (-not $Global:ConnectionStatus.Graph.Connected) {
Show-ErrorMessage -Title “Not Connected” -Message “Please connect to Microsoft Graph first.”
return
}


$searchText = $txtUserSearch.Text.Trim()
if ([string]::IsNullOrWhiteSpace($searchText)) {
    Show-ErrorMessage -Title "Search Empty" -Message "Please enter a search term."
    return
}

try {
    Write-Log "Searching for users with filter: $searchText"
    $cmbUsers.Items.Clear()
    $cmbUsers.Items.Add("Loading...")
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    $users = Get-MgUser -Filter "startswith(userPrincipalName,'$searchText') or startswith(displayName,'$searchText')" -All -ErrorAction Stop | Select-Object UserPrincipalName, DisplayName
    
    $cmbUsers.Items.Clear()
    
    if ($users.Count -eq 0) {
        $cmbUsers.Items.Add("No users found")
        Write-Log "No users found for search: $searchText" -Level WARNING
    }
    else {
        foreach ($user in $users) { $cmbUsers.Items.Add($user.UserPrincipalName) }
        Write-Log "Found $($users.Count) users matching: $searchText"
        $cmbUsers.SelectedIndex = 0
    }
    
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
}
catch {
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    Show-ErrorMessage -Title "Search Failed" -Message $_.Exception.Message
}


})
$tabOffboarding.Controls.Add($btnSearchUser)
$yPosOff += 40

$lblUserDropdown = New-Object System.Windows.Forms.Label
$lblUserDropdown.Location = New-Object System.Drawing.Point(20, $yPosOff)
$lblUserDropdown.Size = New-Object System.Drawing.Size(150, 25)
$lblUserDropdown.Text = “Select User:”
$lblUserDropdown.Font = New-Object System.Drawing.Font(“Segoe UI”, 10, [System.Drawing.FontStyle]::Bold)
$tabOffboarding.Controls.Add($lblUserDropdown)

$cmbUsers = New-Object System.Windows.Forms.ComboBox
$cmbUsers.Location = New-Object System.Drawing.Point(180, $yPosOff)
$cmbUsers.Size = New-Object System.Drawing.Size(510, 25)
$cmbUsers.DropDownStyle = ‘DropDownList’
$cmbUsers.Font = New-Object System.Drawing.Font(“Segoe UI”, 10)
$cmbUsers.Items.Add(“Search for users above…”)
$cmbUsers.SelectedIndex = 0
$tabOffboarding.Controls.Add($cmbUsers)
$yPosOff += 50

$separator1 = New-Object System.Windows.Forms.Label
$separator1.Location = New-Object System.Drawing.Point(20, $yPosOff)
$separator1.Size = New-Object System.Drawing.Size(800, 2)
$separator1.BorderStyle = ‘Fixed3D’
$tabOffboarding.Controls.Add($separator1)
$yPosOff += 20

# Block Sign-In Section

$lblBlockSignIn = New-Object System.Windows.Forms.Label
$lblBlockSignIn.Location = New-Object System.Drawing.Point(20, $yPosOff)
$lblBlockSignIn.Size = New-Object System.Drawing.Size(800, 25)
$lblBlockSignIn.Text = “1. Block Sign-In”
$lblBlockSignIn.Font = New-Object System.Drawing.Font(“Segoe UI”, 11, [System.Drawing.FontStyle]::Bold)
$tabOffboarding.Controls.Add($lblBlockSignIn)
$yPosOff += 30

$btnBlockSignIn = New-Object System.Windows.Forms.Button
$btnBlockSignIn.Location = New-Object System.Drawing.Point(40, $yPosOff)
$btnBlockSignIn.Size = New-Object System.Drawing.Size(200, 30)
$btnBlockSignIn.Text = “Block User Sign-In”
$btnBlockSignIn.BackColor = [System.Drawing.Color]::LightCoral
$btnBlockSignIn.Add_Click({
$selectedUser = $cmbUsers.SelectedItem
if ([string]::IsNullOrWhiteSpace($selectedUser) -or $selectedUser -eq “Search for users above…” -or $selectedUser -eq “No users found” -or $selectedUser -eq “Loading…”) {
Show-ErrorMessage -Title “No User Selected” -Message “Please select a user first.”
return
}


if (-not $Global:ConnectionStatus.Graph.Connected) {
    Show-ErrorMessage -Title "Not Connected" -Message "Please connect to Microsoft Graph first."
    return
}

$confirm = [System.Windows.Forms.MessageBox]::Show(
    "Are you sure you want to BLOCK sign-in for user: $selectedUser?",
    "Confirm Block Sign-In",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Warning
)

if ($confirm -eq 'Yes') {
    try {
        Write-Log "Blocking sign-in for user: $selectedUser"
        Update-MgUser -UserId $selectedUser -AccountEnabled:$false
        Write-Log "Successfully blocked sign-in for: $selectedUser"
        Show-InfoMessage -Title "Success" -Message "Sign-in blocked for $selectedUser"
    }
    catch {
        Show-ErrorMessage -Title "Block Sign-In Failed" -Message $_.Exception.Message
    }
}


})
$tabOffboarding.Controls.Add($btnBlockSignIn)
$yPosOff += 50

# Convert to Shared Mailbox Section

$lblSharedMbx = New-Object System.Windows.Forms.Label
$lblSharedMbx.Location = New-Object System.Drawing.Point(20, $yPosOff)
$lblSharedMbx.Size = New-Object System.Drawing.Size(800, 25)
$lblSharedMbx.Text = “2. Convert to Shared Mailbox”
$lblSharedMbx.Font = New-Object System.Drawing.Font(“Segoe UI”, 11, [System.Drawing.FontStyle]::Bold)
$tabOffboarding.Controls.Add($lblSharedMbx)
$yPosOff += 30

$btnConvertShared = New-Object System.Windows.Forms.Button
$btnConvertShared.Location = New-Object System.Drawing.Point(40, $yPosOff)
$btnConvertShared.Size = New-Object System.Drawing.Size(200, 30)
$btnConvertShared.Text = “Convert to Shared Mailbox”
$btnConvertShared.BackColor = [System.Drawing.Color]::LightBlue
$btnConvertShared.Add_Click({
$selectedUser = $cmbUsers.SelectedItem
if ([string]::IsNullOrWhiteSpace($selectedUser) -or $selectedUser -eq “Search for users above…” -or $selectedUser -eq “No users found” -or $selectedUser -eq “Loading…”) {
Show-ErrorMessage -Title “No User Selected” -Message “Please select a user first.”
return
}


if (-not $Global:ConnectionStatus.Exchange.Connected) {
    Show-ErrorMessage -Title "Not Connected" -Message "Please connect to Exchange Online first."
    return
}

$confirm = [System.Windows.Forms.MessageBox]::Show(
    "Are you sure you want to convert the mailbox to SHARED for user: $selectedUser?",
    "Confirm Convert to Shared",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Warning
)

if ($confirm -eq 'Yes') {
    try {
        Write-Log "Converting mailbox to shared for user: $selectedUser"
        Set-Mailbox -Identity $selectedUser -Type Shared
        Write-Log "Successfully converted mailbox to shared for: $selectedUser"
        Show-InfoMessage -Title "Success" -Message "Mailbox converted to shared for $selectedUser"
    }
    catch {
        Show-ErrorMessage -Title "Convert to Shared Failed" -Message $_.Exception.Message
    }
}


})
$tabOffboarding.Controls.Add($btnConvertShared)
$yPosOff += 50

# Out of Office Section

$lblOOF = New-Object System.Windows.Forms.Label
$lblOOF.Location = New-Object System.Drawing.Point(20, $yPosOff)
$lblOOF.Size = New-Object System.Drawing.Size(800, 25)
$lblOOF.Text = “3. Set Out of Office Reply”
$lblOOF.Font = New-Object System.Drawing.Font(“Segoe UI”, 11, [System.Drawing.FontStyle]::Bold)
$tabOffboarding.Controls.Add($lblOOF)
$yPosOff += 30

$lblOOFMessage = New-Object System.Windows.Forms.Label
$lblOOFMessage.Location = New-Object System.Drawing.Point(40, $yPosOff)
$lblOOFMessage.Size = New-Object System.Drawing.Size(150, 20)
$lblOOFMessage.Text = “OOF Message:”
$lblOOFMessage.Font = New-Object System.Drawing.Font(“Segoe UI”, 9)
$tabOffboarding.Controls.Add($lblOOFMessage)
$yPosOff += 25

$txtOOFMessage = New-Object System.Windows.Forms.TextBox
$txtOOFMessage.Location = New-Object System.Drawing.Point(40, $yPosOff)
$txtOOFMessage.Size = New-Object System.Drawing.Size(650, 80)
$txtOOFMessage.Multiline = $true
$txtOOFMessage.ScrollBars = ‘Vertical’
$txtOOFMessage.Font = New-Object System.Drawing.Font(“Segoe UI”, 9)
$txtOOFMessage.Text = “Thank you for your email. [Employee Name] is no longer with [Company Name]. For assistance, please contact [Alternative Contact].”
$tabOffboarding.Controls.Add($txtOOFMessage)
$yPosOff += 90

$btnSetOOF = New-Object System.Windows.Forms.Button
$btnSetOOF.Location = New-Object System.Drawing.Point(40, $yPosOff)
$btnSetOOF.Size = New-Object System.Drawing.Size(200, 30)
$btnSetOOF.Text = “Set Out of Office”
$btnSetOOF.BackColor = [System.Drawing.Color]::LightYellow
$btnSetOOF.Add_Click({
$selectedUser = $cmbUsers.SelectedItem
if ([string]::IsNullOrWhiteSpace($selectedUser) -or $selectedUser -eq “Search for users above…” -or $selectedUser -eq “No users found” -or $selectedUser -eq “Loading…”) {
Show-ErrorMessage -Title “No User Selected” -Message “Please select a user first.”
return
}


if ([string]::IsNullOrWhiteSpace($txtOOFMessage.Text)) {
    Show-ErrorMessage -Title "No Message" -Message "Please enter an Out of Office message."
    return
}

if (-not $Global:ConnectionStatus.Exchange.Connected) {
    Show-ErrorMessage -Title "Not Connected" -Message "Please connect to Exchange Online first."
    return
}

try {
    Write-Log "Setting OOF for user: $selectedUser"
    $oofMessage = $txtOOFMessage.Text
    Set-MailboxAutoReplyConfiguration -Identity $selectedUser -AutoReplyState Enabled -InternalMessage $oofMessage -ExternalMessage $oofMessage
    Write-Log "Successfully set OOF for: $selectedUser"
    Show-InfoMessage -Title "Success" -Message "Out of Office reply set for $selectedUser"
}
catch {
    Show-ErrorMessage -Title "Set OOF Failed" -Message $_.Exception.Message
}


})
$tabOffboarding.Controls.Add($btnSetOOF)
$yPosOff += 50

# License Management Section

$lblLicenses = New-Object System.Windows.Forms.Label
$lblLicenses.Location = New-Object System.Drawing.Point(20, $yPosOff)
$lblLicenses.Size = New-Object System.Drawing.Size(800, 25)
$lblLicenses.Text = “4. License Management”
$lblLicenses.Font = New-Object System.Drawing.Font(“Segoe UI”, 11, [System.Drawing.FontStyle]::Bold)
$tabOffboarding.Controls.Add($lblLicenses)
$yPosOff += 30

$btnViewLicenses = New-Object System.Windows.Forms.Button
$btnViewLicenses.Location = New-Object System.Drawing.Point(40, $yPosOff)
$btnViewLicenses.Size = New-Object System.Drawing.Size(200, 30)
$btnViewLicenses.Text = “View Current Licenses”
$btnViewLicenses.Add_Click({
$selectedUser = $cmbUsers.SelectedItem
if ([string]::IsNullOrWhiteSpace($selectedUser) -or $selectedUser -eq “Search for users above…” -or $selectedUser -eq “No users found” -or $selectedUser -eq “Loading…”) {
Show-ErrorMessage -Title “No User Selected” -Message “Please select a user first.”
return
}


if (-not $Global:ConnectionStatus.Graph.Connected) {
    Show-ErrorMessage -Title "Not Connected" -Message "Please connect to Microsoft Graph first."
    return
}

try {
    Write-Log "Retrieving licenses for user: $selectedUser"
    $lstLicenses.Items.Clear()
    $lstLicenses.Items.Add("Loading...")
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    $user = Get-MgUser -UserId $selectedUser -Property AssignedLicenses, LicenseAssignmentStates
    $licenses = $user.AssignedLicenses
    
    $lstLicenses.Items.Clear()
    
    if ($licenses.Count -eq 0) {
        $lstLicenses.Items.Add("No licenses assigned")
        Write-Log "No licenses found for user: $selectedUser" -Level WARNING
    }
    else {
        foreach ($license in $licenses) {
            $skuId = $license.SkuId
            $licenseName = $Global:LicenseNameMap[$skuId]
            if ([string]::IsNullOrWhiteSpace($licenseName)) { $licenseName = "Unknown License" }
            $displayText = "$skuId`t$licenseName"
            $lstLicenses.Items.Add($displayText)
        }
        Write-Log "Retrieved $($licenses.Count) licenses for user: $selectedUser"
    }
    
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
}
catch {
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    Show-ErrorMessage -Title "View Licenses Failed" -Message $_.Exception.Message
}


})
$tabOffboarding.Controls.Add($btnViewLicenses)
$yPosOff += 40

$lblLicenseList = New-Object System.Windows.Forms.Label
$lblLicenseList.Location = New-Object System.Drawing.Point(40, $yPosOff)
$lblLicenseList.Size = New-Object System.Drawing.Size(200, 20)
$lblLicenseList.Text = “Assigned Licenses (select to remove):”
$lblLicenseList.Font = New-Object System.Drawing.Font(“Segoe UI”, 9)
$tabOffboarding.Controls.Add($lblLicenseList)
$yPosOff += 25

$lstLicenses = New-Object System.Windows.Forms.ListBox
$lstLicenses.Location = New-Object System.Drawing.Point(40, $yPosOff)
$lstLicenses.Size = New-Object System.Drawing.Size(650, 80)
$lstLicenses.SelectionMode = ‘MultiExtended’
$lstLicenses.Font = New-Object System.Drawing.Font(“Consolas”, 9)
$tabOffboarding.Controls.Add($lstLicenses)
$yPosOff += 90

$btnRemoveLicense = New-Object System.Windows.Forms.Button
$btnRemoveLicense.Location = New-Object System.Drawing.Point(40, $yPosOff)
$btnRemoveLicense.Size = New-Object System.Drawing.Size(200, 30)
$btnRemoveLicense.Text = “Remove Selected License(s)”
$btnRemoveLicense.BackColor = [System.Drawing.Color]::LightSalmon
$btnRemoveLicense.Add_Click({
$selectedUser = $cmbUsers.SelectedItem
if ([string]::IsNullOrWhiteSpace($selectedUser) -or $selectedUser -eq “Search for users above…” -or $selectedUser -eq “No users found” -or $selectedUser -eq “Loading…”) {
Show-ErrorMessage -Title “No User Selected” -Message “Please select a user first.”
return
}


if ($lstLicenses.SelectedItems.Count -eq 0) {
    Show-ErrorMessage -Title "No License Selected" -Message "Please select at least one license to remove."
    return
}

if (-not $Global:ConnectionStatus.Graph.Connected) {
    Show-ErrorMessage -Title "Not Connected" -Message "Please connect to Microsoft Graph first."
    return
}

$licensesToRemove = @()
foreach ($item in $lstLicenses.SelectedItems) {
    $skuId = $item.ToString().Split("`t")[0]
    $licensesToRemove += $skuId
}

$confirm = [System.Windows.Forms.MessageBox]::Show(
    "Are you sure you want to REMOVE the following license(s) from $selectedUser`?`n`n$($licensesToRemove -join "`n")",
    "Confirm License Removal",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Warning
)

if ($confirm -eq 'Yes') {
    try {
        Write-Log "Removing licenses for user: $selectedUser - Licenses: $($licensesToRemove -join ', ')"
        Set-MgUserLicense -UserId $selectedUser -AddLicenses @() -RemoveLicenses $licensesToRemove
        Write-Log "Successfully removed licenses for: $selectedUser"
        Show-InfoMessage -Title "Success" -Message "License(s) removed for $selectedUser"
        
        # Refresh license list
        $lstLicenses.Items.Clear()
        $user = Get-MgUser -UserId $selectedUser -Property AssignedLicenses
        $licenses = $user.AssignedLicenses
        
        if ($licenses.Count -eq 0) {
            $lstLicenses.Items.Add("No licenses assigned")
        }
        else {
            foreach ($license in $licenses) {
                $skuId = $license.SkuId
                $licenseName = $Global:LicenseNameMap[$skuId]
                if ([string]::IsNullOrWhiteSpace($licenseName)) { $licenseName = "Unknown License" }
                $displayText = "$skuId`t$licenseName"
                $lstLicenses.Items.Add($displayText)
            }
        }
    }
    catch {
        Show-ErrorMessage -Title "Remove License Failed" -Message $_.Exception.Message
    }
}


})
$tabOffboarding.Controls.Add($btnRemoveLicense)

# Update Connection Status Function

function Update-ConnectionStatus {
foreach ($key in $Global:StatusLabels.Keys) {
$status = $Global:ConnectionStatus[$key]
$label = $Global:StatusLabels[$key]


    if ($status.Connected) {
        $label.Text = "Status: Connected | Tenant: $($status.Tenant)"
        $label.ForeColor = [System.Drawing.Color]::Green
    }
    else {
        $label.Text = "Status: Not Connected | Tenant: N/A"
        $label.ForeColor = [System.Drawing.Color]::Red
    }
}

$lastCheck = ($Global:ConnectionStatus.Values | Where-Object { $_.LastCheck } | Sort-Object -Property LastCheck -Descending | Select-Object -First 1).LastCheck
if ($lastCheck) {
    $Global:lblLastCheck.Text = "Last status check: $(Get-Date $lastCheck -Format 'HH:mm:ss') | Auto-refresh every 30 seconds"
}

if ($Global:lblCurrentScopes) {
    $Global:lblCurrentScopes.Text = "Current Scopes: $($Global:GraphScopes -join ', ')"
}


}

# Initial status check

Test-GraphConnection | Out-Null
Test-ExchangeConnection | Out-Null
Test-MSOnlineConnection | Out-Null
Test-AzureADConnection | Out-Null
Test-SharePointConnection | Out-Null
Test-TeamsConnection | Out-Null
Update-ConnectionStatus

# Show the form

Write-Log "=== Ultimate Microsoft 365 Management Tool v1.8 Started ==="
Write-Log "Log file location: $Global:LogFile"

try {
    [void]$form.ShowDialog()
}
catch {
    Write-Log "Error in main form: $_" -Level "ERROR"
    Show-ErrorMessage -Title "Application Error" -Message "An error occurred: $_"
}
finally {
    # Cleanup on exit
    if ($statusTimer) {
        try {
            $statusTimer.Stop()
            $statusTimer.Dispose()
        }
        catch {
            Write-Log "Error cleaning up timer: $_" -Level "WARNING"
        }
    }
    Write-Log "=== Ultimate Microsoft 365 Management Tool v1.8 Closed ==="
}
