# Mailbox Permissions WPF GUI
# Requires: Exchange Online PowerShell module and Microsoft Graph PowerShell module

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Xaml


Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

$console = [Console.Window]::GetConsoleWindow()

# 0 hide
[Console.Window]::ShowWindow($console, 0) | Out-Null


# Icon is set directly in XAML - no function needed

# Embedded XAML
$XAML = @"
  <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
          xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
          MinWidth="1000"
          MinHeight="700"
          Width="1100"
          Height="811"
          Background="#F8F9FA"
          Icon="https://axcientrestore.blob.core.windows.net/win11/SOS.ico"
          Title="Mailbox Permissions Manager"
          WindowStartupLocation="CenterScreen"
          ResizeMode="CanResize">
  <Window.Resources>
    <!-- Standard Button Style -->
    <Style x:Key="StandardButton" TargetType="Button">
      <Setter Property="Height" Value="32" />
      <Setter Property="MinWidth" Value="100" />
      <Setter Property="Padding" Value="12,6" />
      <Setter Property="Margin" Value="2" />
      <Setter Property="FontSize" Value="12" />
      <Setter Property="FontWeight" Value="SemiBold" />
      <Setter Property="Cursor" Value="Hand" />
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="1"
                    CornerRadius="4">
              <ContentPresenter Margin="{TemplateBinding Padding}"
                                HorizontalAlignment="Center"
                                VerticalAlignment="Center" />
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Opacity" Value="0.8" />
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter Property="Opacity" Value="0.6" />
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Background" Value="#E9ECEF" />
                <Setter Property="Foreground" Value="#6C757D" />
                <Setter Property="Opacity" Value="0.7" />
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <!-- Standard TextBox Style -->
    <Style x:Key="StandardTextBox" TargetType="TextBox">
      <Setter Property="Height" Value="42" />
      <Setter Property="Padding" Value="8,10" />
      <Setter Property="Margin" Value="2" />
      <Setter Property="FontSize" Value="12" />
      <Setter Property="Foreground" Value="Black" />
      <Setter Property="Background" Value="White" />
      <Setter Property="VerticalContentAlignment" Value="Center" />
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="#CCCCCC"
                    BorderThickness="1"
                    CornerRadius="4">
              <ScrollViewer x:Name="PART_ContentHost"
                            Margin="0"
                            VerticalAlignment="Stretch"
                            CanContentScroll="False"
                            Foreground="{TemplateBinding Foreground}"
                            HorizontalScrollBarVisibility="Disabled"
                            VerticalScrollBarVisibility="Disabled" />
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <!-- Standard Label Style -->
    <Style x:Key="StandardLabel" TargetType="Label">
      <Setter Property="FontSize" Value="12" />
      <Setter Property="FontWeight" Value="SemiBold" />
      <Setter Property="VerticalAlignment" Value="Center" />
      <Setter Property="Margin" Value="0,0,8,0" />
    </Style>
  </Window.Resources>
  <Grid Margin="15">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto" />
      <RowDefinition Height="Auto" />
      <RowDefinition Height="*" />
      <RowDefinition Height="Auto" />
    </Grid.RowDefinitions>
      <!-- Header - Background Image Only -->
      <Grid Margin="0,0,0,15" Grid.Row="0">
        <Image Height="80"
               Margin="0,0,0,0"
               HorizontalAlignment="Stretch"
               VerticalAlignment="Center"
               Source="https://axcientrestore.blob.core.windows.net/win11/SOS-Banner4.png"
               Stretch="Uniform"
               Grid.Column="0"
               Grid.Row="0" />
      </Grid>
    <!-- Connection Section -->
    <GroupBox Margin="0,0,0,10"
              FontWeight="SemiBold"
              Header="Connection Settings"
              Padding="15"
              Grid.Row="1">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto" />
          <ColumnDefinition Width="*" />
          <ColumnDefinition Width="Auto" />
          <ColumnDefinition Width="Auto" />
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto" />
          <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        <Label Content="Activity:"
               Margin="0,0,8,0"
               HorizontalAlignment="Stretch"
               VerticalAlignment="Center"
               Style="{StaticResource StandardLabel}"
               Grid.Column="0"
               Grid.Row="0" />
        <Button Name="btnConnect"
                Content="Connect"
                Background="#28A745"
                Foreground="White"
                Style="{StaticResource StandardButton}"
                Grid.Column="2"
                Grid.Row="0" />
        <Label Content="Status:"
               Style="{StaticResource StandardLabel}"
               Grid.Column="0"
               Grid.Row="1" />
        <TextBlock Name="lblConnectionStatus"
                   VerticalAlignment="Center"
                   FontSize="12"
                   FontWeight="SemiBold"
                   Foreground="Red"
                   Text="Not Connected"
                   Grid.Column="1"
                   Grid.Row="1" />
        <Button x:Name="btnDisconnect"
                Content="Disconnect"
                Background="#DC3545"
                Foreground="White"
                IsEnabled="False"
                Style="{StaticResource StandardButton}"
                Grid.Column="2"
                Grid.Row="1" />
        <Grid x:Name="activityBar"
              Height="24"
              Margin="0,0,0,0"
              HorizontalAlignment="Stretch"
              VerticalAlignment="Center"
              Grid.Column="1"
              Grid.Row="0">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*" />
          </Grid.ColumnDefinitions>
          <ProgressBar x:Name="ProgressBar_Copy1"
                       Height="24"
                       Margin="0,0,6.40000000000009,0"
                       HorizontalAlignment="Stretch"
                       VerticalAlignment="Center"
                       Foreground="#007BFF"
                       IsIndeterminate="False"
                       Value="0"
                       Visibility="Visible"
                       Grid.Column="0"
                       Grid.Row="0" />
          <TextBlock x:Name="ProgressTextBlock_Copy1"
                     HorizontalAlignment="Center"
                     VerticalAlignment="Center"
                     FontSize="11"
                     FontWeight="SemiBold"
                     Foreground="Black"
                     Text=""
                     Grid.Column="0"
                     Grid.Row="0" />
        </Grid>
      </Grid>
    </GroupBox>
    <!-- Main Content -->
    <Grid Grid.Row="2">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="2*" />
        <ColumnDefinition Width="*" />
      </Grid.ColumnDefinitions>
      <!-- Left Panel - Mailbox Operations -->
      <GroupBox Margin="0,0,10,0"
                HorizontalAlignment="Stretch"
                VerticalAlignment="Stretch"
                FontWeight="SemiBold"
                Header="Mailbox Operations"
                Padding="15"
                Grid.Column="0"
                Grid.Row="0">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
          </Grid.RowDefinitions>
          <!-- Mailbox and Trustee Input -->
          <Grid Margin="0,0,0,15" Grid.Row="0">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto" />
              <ColumnDefinition Width="*" />
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto" />
              <RowDefinition Height="Auto" />
            </Grid.RowDefinitions>
            <Label Content="Mailbox Identity:"
                   Style="{StaticResource StandardLabel}"
                   Grid.Column="0"
                   Grid.Row="0" />
            <TextBox Name="txtMailbox"
                     Height="37"
                     Style="{StaticResource StandardTextBox}"
                     Text="Enter target UPN here"
                     Grid.Column="1"
                     Grid.Row="0" />
            <Label Content="Trustee:"
                   Style="{StaticResource StandardLabel}"
                   Grid.Column="0"
                   Grid.Row="1" />
            <TextBox Name="txtTrustee"
                     Height="37"
                     Style="{StaticResource StandardTextBox}"
                     Text=""
                     Grid.Column="1"
                     Grid.Row="1" />
          </Grid>
          <!-- Permission Buttons -->
          <Grid Margin="0,0,0,10" Grid.Row="1">
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto" />
              <RowDefinition Height="Auto" />
              <RowDefinition Height="Auto" />
              <RowDefinition Height="Auto" />
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*" />
              <ColumnDefinition Width="*" />
            </Grid.ColumnDefinitions>
            <!-- Action Buttons -->
            <Button Name="btnValidateMailbox"
                    Content="Validate Mailbox"
                    IsEnabled="False"
                    Style="{StaticResource StandardButton}"
                    Grid.Column="0"
                    Grid.Row="0" />
            <Button x:Name="btnListPermissions"
                    Content="List All Permissions"
                    IsEnabled="False"
                    Style="{StaticResource StandardButton}"
                    Grid.Column="1"
                    Grid.Row="0" />
            <!-- SendAs Permissions -->
            <Button Name="btnAddSendAs"
                    Content="Add SendAs"
                    IsEnabled="False"
                    Style="{StaticResource StandardButton}"
                    Grid.Column="0"
                    Grid.Row="1" />
            <Button Name="btnRemoveSendAs"
                    Content="Remove SendAs"
                    IsEnabled="False"
                    Style="{StaticResource StandardButton}"
                    Grid.Column="1"
                    Grid.Row="1" />
            <!-- FullAccess Permissions -->
            <Button Name="btnAddFullAccess"
                    Content="Add FullAccess"
                    IsEnabled="False"
                    Style="{StaticResource StandardButton}"
                    Grid.Column="0"
                    Grid.Row="2" />
            <Button Name="btnRemoveFullAccess"
                    Content="Remove FullAccess"
                    IsEnabled="False"
                    Style="{StaticResource StandardButton}"
                    Grid.Column="1"
                    Grid.Row="2" />
            <!-- SendOnBehalf Permissions -->
            <Button Name="btnAddSendOnBehalf"
                    Content="Add SendOnBehalf"
                    IsEnabled="False"
                    Style="{StaticResource StandardButton}"
                    Grid.Column="0"
                    Grid.Row="3" />
            <Button Name="btnRemoveSendOnBehalf"
                    Content="Remove SendOnBehalf"
                    IsEnabled="False"
                    Style="{StaticResource StandardButton}"
                    Grid.Column="1"
                    Grid.Row="3" />
          </Grid>
          <!-- Calendar Permissions -->
          <GroupBox Margin="0,0,0,15"
                    FontWeight="SemiBold"
                    Header="Calendar Permissions"
                    Padding="15"
                    Grid.Row="2">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto" />
                <RowDefinition Height="Auto" />
                <RowDefinition Height="Auto" />
              </Grid.RowDefinitions>
              <Grid Margin="0,0,0,10" Grid.Row="0">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto" />
                  <ColumnDefinition Width="*" />
                  <ColumnDefinition Width="Auto" />
                </Grid.ColumnDefinitions>
                <Label Content="Calendar Role:"
                       Style="{StaticResource StandardLabel}"
                       Grid.Column="0" />
                <ComboBox Name="cmbCalendarRole"
                          Height="32"
                          Margin="2"
                          FontSize="12"
                          Padding="8,4"
                          SelectedIndex="0"
                          Grid.Column="1">
                  <ComboBoxItem Content="Owner" />
                  <ComboBoxItem Content="PublishingEditor" />
                  <ComboBoxItem Content="Editor" />
                  <ComboBoxItem Content="PublishingAuthor" />
                  <ComboBoxItem Content="Author" />
                  <ComboBoxItem Content="NonEditingAuthor" />
                  <ComboBoxItem Content="Reviewer" />
                  <ComboBoxItem Content="Contributor" />
                  <ComboBoxItem Content="None" />
                </ComboBox>
                <Button Name="btnAddCalendar"
                        Content="Add Calendar"
                        IsEnabled="False"
                        Style="{StaticResource StandardButton}"
                        Grid.Column="2" />
              </Grid>
              <Grid Margin="0,0,0,10" Grid.Row="1">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*" />
                  <ColumnDefinition Width="Auto" />
                </Grid.ColumnDefinitions>
                <TextBox Name="txtCalendarTrustee"
                         Height="37"
                         Style="{StaticResource StandardTextBox}"
                         Text="Enter target UPN here"
                         Grid.Column="0"
                         Grid.Row="0" />
                <Button Name="btnRemoveCalendar"
                        Content="Remove Calendar"
                        IsEnabled="False"
                        Style="{StaticResource StandardButton}"
                        Grid.Column="1" />
              </Grid>
              <Button Name="btnListCalendar"
                      Content="List Calendar Permissions"
                      IsEnabled="False"
                      Style="{StaticResource StandardButton}"
                      Grid.Row="2" />
            </Grid>
          </GroupBox>
        </Grid>
      </GroupBox>
      <!-- Right Panel - Log and Status -->
      <GroupBox Margin="10,0,0,0"
                HorizontalAlignment="Stretch"
                VerticalAlignment="Stretch"
                FontWeight="SemiBold"
                Header="Activity Log"
                Padding="15"
                Grid.Column="1"
                Grid.Row="0">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
          </Grid.RowDefinitions>
          <ScrollViewer HorizontalScrollBarVisibility="Auto"
                        VerticalScrollBarVisibility="Auto"
                        Grid.Row="0">
            <TextBlock Name="txtLog"
                       Background="#F8F9FA"
                       FontFamily="Consolas"
                       FontSize="10"
                       LineHeight="16"
                       Padding="10"
                       Text="Ready to connect..."
                       TextWrapping="Wrap" />
          </ScrollViewer>
          <Button Name="btnClearLog"
                  Content="Clear Log"
                  IsEnabled="False"
                  Style="{StaticResource StandardButton}"
                  Grid.Row="1" />
        </Grid>
      </GroupBox>
    </Grid>
    <!-- Status Bar -->
    <StatusBar Height="28"
               Margin="0,15,0,0"
               Background="#E9ECEF"
               Grid.Row="3">
      <StatusBarItem>
        <TextBlock Name="lblStatus"
                   FontSize="12"
                   FontWeight="SemiBold"
                   Foreground="Black"
                   Text="Ready" />
      </StatusBarItem>
      <Separator />
      <StatusBarItem>
        <TextBlock Name="lblUser"
                   FontSize="12"
                   FontWeight="SemiBold"
                   Foreground="Black"
                   Text="User: Not Connected" />
      </StatusBarItem>
    </StatusBar>
  </Grid>
</Window>
"@

# Icon is set directly in XAML - no additional processing needed

# Load the XAML
[xml]$XAMLDoc = $XAML
$XAMLReader = [System.Xml.XmlNodeReader]::new($XAMLDoc)
$Window = [Windows.Markup.XamlReader]::Load($XAMLReader)

# Icon is already set in XAML

# Get references to controls
$txtMailbox = $Window.FindName("txtMailbox")
$txtTrustee = $Window.FindName("txtTrustee")
$txtCalendarTrustee = $Window.FindName("txtCalendarTrustee")
$cmbCalendarRole = $Window.FindName("cmbCalendarRole")
$txtLog = $Window.FindName("txtLog")
$lblConnectionStatus = $Window.FindName("lblConnectionStatus")
$lblStatus = $Window.FindName("lblStatus")
$lblUser = $Window.FindName("lblUser")

# Activity bar controls
$activityBar = $Window.FindName("activityBar")
$progressBar = $Window.FindName("ProgressBar_Copy1")
$progressTextBlock = $Window.FindName("ProgressTextBlock_Copy1")

# Get ScrollViewer reference for log scrolling
$logScrollViewer = $txtLog.Parent

# Set default activity bar message
$progressTextBlock.Text = "Waiting for command..."

# Button references
$btnConnect = $Window.FindName("btnConnect")
$btnRefreshTenant = $Window.FindName("btnRefreshTenant")
$btnDisconnect = $Window.FindName("btnDisconnect")
$btnAddSendAs = $Window.FindName("btnAddSendAs")
$btnRemoveSendAs = $Window.FindName("btnRemoveSendAs")
$btnAddFullAccess = $Window.FindName("btnAddFullAccess")
$btnRemoveFullAccess = $Window.FindName("btnRemoveFullAccess")
$btnAddSendOnBehalf = $Window.FindName("btnAddSendOnBehalf")
$btnRemoveSendOnBehalf = $Window.FindName("btnRemoveSendOnBehalf")
$btnAddCalendar = $Window.FindName("btnAddCalendar")
$btnRemoveCalendar = $Window.FindName("btnRemoveCalendar")
$btnListCalendar = $Window.FindName("btnListCalendar")
$btnListPermissions = $Window.FindName("btnListPermissions")
$btnValidateMailbox = $Window.FindName("btnValidateMailbox")
$btnClearLog = $Window.FindName("btnClearLog")


# Global variables
$Global:ExchangeConnected = $false
$Global:GraphConnected = $false
$Global:Tenant = ""
$Global:LogPath = "C:\Temp\Powershell-Logging"
$Global:LogFile = "$LogPath\MailboxPermissionLog.txt"

# Ensure log directory exists
if (!(Test-Path $Global:LogPath)) { 
    New-Item -ItemType Directory -Path $Global:LogPath -Force | Out-Null 
}

# Logging function
function Write-Log {
    param([string]$Message)
    $TimeStamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $CurrentUser = "$($env:USERDOMAIN)\$($env:USERNAME)"
    Add-Content -Path $Global:LogFile -Value "$TimeStamp - $CurrentUser - $Message"
}

# Update log display
function Update-Log {
    param([string]$Message, [string]$Color = "Black")
    $TimeStamp = (Get-Date).ToString("HH:mm:ss")
    $LogEntry = "[$TimeStamp] $Message"
    
    $Window.Dispatcher.Invoke([Action]{
        $txtLog.Text += "`n$LogEntry"
        if ($logScrollViewer) {
            $logScrollViewer.ScrollToEnd()
        }
    })
    
    Write-Log $Message
}

# Update status
function Update-Status {
    param([string]$Message)
    $Window.Dispatcher.Invoke([Action]{
        $lblStatus.Text = $Message
    })
}

# Start activity bar animation
function Start-ActivityBar {
    param([string]$Message = "Processing...")
    $Window.Dispatcher.Invoke([Action]{
        $progressTextBlock.Text = $Message
        $progressBar.IsIndeterminate = $true
        $progressBar.Value = 0
    })
}

# Stop activity bar animation
function Stop-ActivityBar {
    $Window.Dispatcher.Invoke([Action]{
        $progressBar.IsIndeterminate = $false
        $progressBar.Value = 0
        $progressTextBlock.Text = "Waiting for command..."
    })
}

# Enable/disable controls based on connection status
function Set-ControlsEnabled {
    param([bool]$Enabled)
    $Window.Dispatcher.Invoke([Action]{
        # Set enabled state
        $btnAddSendAs.IsEnabled = $Enabled
        $btnRemoveSendAs.IsEnabled = $Enabled
        $btnAddFullAccess.IsEnabled = $Enabled
        $btnRemoveFullAccess.IsEnabled = $Enabled
        $btnAddSendOnBehalf.IsEnabled = $Enabled
        $btnRemoveSendOnBehalf.IsEnabled = $Enabled
        $btnAddCalendar.IsEnabled = $Enabled
        $btnRemoveCalendar.IsEnabled = $Enabled
        $btnListCalendar.IsEnabled = $Enabled
        $btnListPermissions.IsEnabled = $Enabled
        $btnValidateMailbox.IsEnabled = $Enabled
        $btnDisconnect.IsEnabled = $Enabled
        $btnClearLog.IsEnabled = $Enabled
        if ($btnRefreshTenant) {
            $btnRefreshTenant.IsEnabled = $Enabled
        }
        
        # Set colors when enabling buttons
        if ($Enabled) {
            $btnAddSendAs.Background = "#007BFF"
            $btnAddSendAs.Foreground = "White"
            $btnRemoveSendAs.Background = "#DC3545"
            $btnRemoveSendAs.Foreground = "White"
            $btnAddFullAccess.Background = "#007BFF"
            $btnAddFullAccess.Foreground = "White"
            $btnRemoveFullAccess.Background = "#DC3545"
            $btnRemoveFullAccess.Foreground = "White"
            $btnAddSendOnBehalf.Background = "#007BFF"
            $btnAddSendOnBehalf.Foreground = "White"
            $btnRemoveSendOnBehalf.Background = "#DC3545"
            $btnRemoveSendOnBehalf.Foreground = "White"
            $btnAddCalendar.Background = "#28A745"
            $btnAddCalendar.Foreground = "White"
            $btnRemoveCalendar.Background = "#DC3545"
            $btnRemoveCalendar.Foreground = "White"
            $btnListCalendar.Background = "#6C757D"
            $btnListCalendar.Foreground = "White"
            $btnListPermissions.Background = "#343A40"
            $btnListPermissions.Foreground = "White"
            $btnValidateMailbox.Background = "#343A40"
            $btnValidateMailbox.Foreground = "White"
            $btnClearLog.Background = "#6C757D"
            $btnClearLog.Foreground = "White"
        }
        
        #Update-Log "Set-ControlsEnabled called with: $Enabled" "Blue"
    })
}

# Validate mailbox exists
function Test-MailboxExists {
    param([string]$Mailbox)
    try {
        Get-Mailbox -Identity $Mailbox -ErrorAction Stop | Out-Null
        return $true
    } catch {
        Update-Log "Mailbox [$Mailbox] does not exist or cannot be found." "Red"
        return $false
    }
}

# Connect to Exchange Online
function Connect-Exchange {
    # Use default tenant - will be determined during authentication
    $Tenant = "common"  # Use common endpoint to allow any tenant
    
    # Check if Exchange Online module is available
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Update-Log "Exchange Online PowerShell module not found. Please install it first:" "Red"
        Update-Log "Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber" "Red"
        return $false
    }
    
    try {
        # Start activity bar and update UI immediately
        Start-ActivityBar "Connecting to Exchange Online..."
        Update-Status "Connecting to Exchange Online..."
        Update-Log "Attempting to connect to Exchange Online..."
        Update-Log "This may open a browser window for authentication..." "Blue"
        
        # Force UI update before blocking operation
        $Window.Dispatcher.Invoke([Action]{})
        
        # Import the module if not already loaded
        if (-not (Get-Module -Name ExchangeOnlineManagement)) {
            Import-Module ExchangeOnlineManagement -Force
            Update-Log "Exchange Online module imported successfully" "Blue"
        }
        
        # Update activity bar message before authentication
        Start-ActivityBar "Authenticating with Microsoft 365..."
        $Window.Dispatcher.Invoke([Action]{})
        
        # Small delay to ensure UI updates are visible
        Start-Sleep -Milliseconds 500
        
        # Update log immediately and force UI refresh
        $TimeStamp = (Get-Date).ToString("HH:mm:ss")
        $LogEntry = "[$TimeStamp] Connecting to M365..."
        $txtLog.Text += "`n$LogEntry"
        if ($logScrollViewer) {
            $logScrollViewer.ScrollToEnd()
        }
        # Write to log file
        Write-Log "Connecting to M365..."
        # Force immediate UI update with high priority
        $Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Send)
        # Process any pending UI updates
        [System.Windows.Forms.Application]::DoEvents()
        # Small delay to ensure message is visible
        Start-Sleep -Milliseconds 300
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $Global:ExchangeConnected = $true
        
        # Get the actual connected tenant information
        $ActualTenant = $null
        try {
            $ConnectionInfo = Get-ConnectionInformation
            if ($ConnectionInfo -and $ConnectionInfo.Organization -and $ConnectionInfo.Organization -ne "Enter tenant value" -and $ConnectionInfo.Organization -ne "common") {
                $ActualTenant = $ConnectionInfo.Organization
                Update-Log "Actual connected tenant: $ActualTenant" "Blue"
            } else {
                # Fallback: Try to get tenant from current session
                try {
                    $SessionInfo = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" }
                    if ($SessionInfo -and $SessionInfo.ComputerName) {
                        $ActualTenant = $SessionInfo.ComputerName -replace "outlook.office365.com", "" -replace "^\.", ""
                        if ($ActualTenant) {
                            Update-Log "Retrieved tenant from session: $ActualTenant" "Blue"
                        }
                    }
                } catch {
                    Update-Log "Could not get tenant from session: $($_.Exception.Message)" "Red"
                }
            }
        } catch {
            Update-Log "Could not get connection information: $($_.Exception.Message)" "Red"
        }
        
        # If we still don't have a tenant, try to extract from authenticated user
        if (-not $ActualTenant) {
            try {
                $AuthUser = Get-ConnectionInformation | Select-Object -ExpandProperty UserPrincipalName -ErrorAction SilentlyContinue
                if ($AuthUser -and $AuthUser -like "*@*") {
                    $ActualTenant = $AuthUser.Split('@')[1]
                    Update-Log "Extracted tenant from user email: $ActualTenant" "Blue"
                }
            } catch {
                Update-Log "Could not extract tenant from user email" "Red"
            }
        }
        
        # Final fallback
        if (-not $ActualTenant) {
            $ActualTenant = "Unknown Tenant"
            Update-Log "Could not determine tenant name" "Red"
        }
        
        # Update global tenant variable with actual tenant
        $Global:Tenant = $ActualTenant
        
        # Get the authenticated user's email address
        try {
            $AuthUser = Get-ConnectionInformation | Select-Object -ExpandProperty UserPrincipalName -ErrorAction SilentlyContinue
            if (-not $AuthUser) {
                # Fallback: try to get from current context
                $AuthUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                if ($AuthUser -like "*\*") {
                    $AuthUser = $AuthUser.Split('\')[1]
                }
            }
            Update-Log "Authenticated user: $AuthUser" "Blue"
        } catch {
            $AuthUser = "user@$ActualTenant"
            Update-Log "Could not determine authenticated user, using default: $AuthUser" "Blue"
        }
        
        $Window.Dispatcher.Invoke([Action]{
            $lblConnectionStatus.Text = "Connected to $ActualTenant"
            $lblConnectionStatus.Foreground = "Green"
            $lblUser.Text = "User: $($env:USERDOMAIN)\$($env:USERNAME)"
            $txtTrustee.Text = $AuthUser
        })
        
        # Update global tenant variable
        $Global:Tenant = $ActualTenant
        
        Set-ControlsEnabled $true
        Stop-ActivityBar
        Update-Status "Connected to Exchange Online"
        Update-Log "Successfully connected to Exchange Online for tenant: $ActualTenant" "Green"
        return $true
    } catch {
        Stop-ActivityBar
        $Global:ExchangeConnected = $false
        $Window.Dispatcher.Invoke([Action]{
            $lblConnectionStatus.Text = "Connection Failed"
            $lblConnectionStatus.Foreground = "Red"
        })
        Update-Status "Connection failed"
        Update-Log "Failed to connect to Exchange Online: $($_.Exception.Message)" "Red"
        Update-Log "Error details: $($_.Exception.GetType().FullName)" "Red"
        if ($_.Exception.InnerException) {
            Update-Log "Inner exception: $($_.Exception.InnerException.Message)" "Red"
        }
        return $false
    }
}

# Refresh tenant information
function Refresh-TenantInfo {
    if ($Global:ExchangeConnected) {
        try {
            Start-ActivityBar "Refreshing tenant information..."
            $ActualTenant = $null
            
            # Try Get-ConnectionInformation first
            try {
                $ConnectionInfo = Get-ConnectionInformation
                if ($ConnectionInfo -and $ConnectionInfo.Organization -and $ConnectionInfo.Organization -ne "Enter tenant value" -and $ConnectionInfo.Organization -ne "common") {
                    $ActualTenant = $ConnectionInfo.Organization
                    Update-Log "Refreshed tenant from connection info: $ActualTenant" "Blue"
                } else {
                    # Fallback: Try to get tenant from current session
                    try {
                        $SessionInfo = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" }
                        if ($SessionInfo -and $SessionInfo.ComputerName) {
                            $ActualTenant = $SessionInfo.ComputerName -replace "outlook.office365.com", "" -replace "^\.", ""
                            if ($ActualTenant) {
                                Update-Log "Refreshed tenant from session: $ActualTenant" "Blue"
                            }
                        }
                    } catch {
                        Update-Log "Could not get tenant from session: $($_.Exception.Message)" "Red"
                    }
                }
            } catch {
                Update-Log "Could not get connection information: $($_.Exception.Message)" "Red"
            }
            
            # If we still don't have a tenant, try to extract from authenticated user
            if (-not $ActualTenant) {
                try {
                    $AuthUser = Get-ConnectionInformation | Select-Object -ExpandProperty UserPrincipalName -ErrorAction SilentlyContinue
                    if ($AuthUser -and $AuthUser -like "*@*") {
                        $ActualTenant = $AuthUser.Split('@')[1]
                        Update-Log "Extracted tenant from user email: $ActualTenant" "Blue"
                    }
                } catch {
                    Update-Log "Could not extract tenant from user email" "Red"
                }
            }
            
            # Final fallback - use the global tenant variable if it exists
            if (-not $ActualTenant -and $Global:Tenant) {
                $ActualTenant = $Global:Tenant
                Update-Log "Using cached tenant: $ActualTenant" "Blue"
            }
            
            if ($ActualTenant) {
                # Get the authenticated user's email address
                try {
                    $AuthUser = Get-ConnectionInformation | Select-Object -ExpandProperty UserPrincipalName -ErrorAction SilentlyContinue
                    if (-not $AuthUser) {
                        $AuthUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                        if ($AuthUser -like "*\*") {
                            $AuthUser = $AuthUser.Split('\')[1]
                        }
                    }
                } catch {
                    $AuthUser = "user@$ActualTenant"
                }
                
                $Window.Dispatcher.Invoke([Action]{
                    $lblConnectionStatus.Text = "Connected to $ActualTenant"
                    $txtTrustee.Text = $AuthUser
                })
                $Global:Tenant = $ActualTenant
                Stop-ActivityBar
                Update-Log "Successfully refreshed tenant info: $ActualTenant" "Green"
                Update-Log "Updated trustee to authenticated user: $AuthUser" "Blue"
            } else {
                Stop-ActivityBar
                Update-Log "Could not determine actual tenant name - all detection methods failed" "Red"
                Update-Log "Connection may be in an inconsistent state. Try disconnecting and reconnecting." "Yellow"
            }
        } catch {
            Stop-ActivityBar
            Update-Log "Could not refresh tenant information: $($_.Exception.Message)" "Red"
        }
    } else {
        Update-Log "Not connected to Exchange Online" "Red"
    }
}

# Disconnect from Exchange Online
function Disconnect-Exchange {
    try {
        Start-ActivityBar "Disconnecting from all services..."
        Update-Status "Disconnecting from all services..."
        Update-Log "Starting comprehensive disconnect from all Microsoft 365 services..." "Blue"
        
        # Disconnect from Exchange Online
        if ($Global:ExchangeConnected) {
            try {
                Update-Log "Disconnecting from Exchange Online..." "Blue"
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
                Update-Log "Successfully disconnected from Exchange Online" "Green"
            } catch {
                Update-Log "Warning: Error disconnecting from Exchange Online: $($_.Exception.Message)" "Yellow"
            }
        }
        
        # Disconnect from Microsoft Graph
        if ($Global:GraphConnected) {
            try {
                Update-Log "Disconnecting from Microsoft Graph..." "Blue"
                Disconnect-MgGraph -ErrorAction Stop
                Update-Log "Successfully disconnected from Microsoft Graph" "Green"
            } catch {
                Update-Log "Warning: Error disconnecting from Microsoft Graph: $($_.Exception.Message)" "Yellow"
            }
        }
        
        # Clear all PowerShell sessions
        try {
            Update-Log "Clearing all PowerShell sessions..." "Blue"
            $ExchangeSessions = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" }
            if ($ExchangeSessions) {
                Remove-PSSession $ExchangeSessions -ErrorAction SilentlyContinue
                Update-Log "Removed $($ExchangeSessions.Count) Exchange PowerShell sessions" "Green"
            }
            
            $GraphSessions = Get-PSSession | Where-Object { $_.ConfigurationName -like "*Graph*" }
            if ($GraphSessions) {
                Remove-PSSession $GraphSessions -ErrorAction SilentlyContinue
                Update-Log "Removed $($GraphSessions.Count) Graph PowerShell sessions" "Green"
            }
        } catch {
            Update-Log "Warning: Error clearing PowerShell sessions: $($_.Exception.Message)" "Yellow"
        }
        
        # Clear connection information cache
        try {
            Update-Log "Clearing connection information cache..." "Blue"
            # Clear any cached connection info
            if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
                try {
                    $ConnectionInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
                    if ($ConnectionInfo) {
                        Update-Log "Cleared connection information cache" "Green"
                    }
                } catch {
                    # Connection info already cleared
                }
            }
        } catch {
            Update-Log "Warning: Error clearing connection cache: $($_.Exception.Message)" "Yellow"
        }
        
        # Reset global variables
        $Global:ExchangeConnected = $false
        $Global:GraphConnected = $false
        $Global:Tenant = ""
        
        $Window.Dispatcher.Invoke([Action]{
            $lblConnectionStatus.Text = "Not Connected"
            $lblConnectionStatus.Foreground = "Red"
            $lblUser.Text = "User: Not Connected"
        })
        
        Set-ControlsEnabled $false
        Stop-ActivityBar
        Update-Status "Disconnected"
        Update-Log "Disconnected from all Microsoft 365 services" "Green"
        
    } catch {
        Stop-ActivityBar
        Update-Log "Error during disconnect: $($_.Exception.Message)" "Red"
        # Force reset connection state even if disconnect fails
        $Global:ExchangeConnected = $false
        $Global:GraphConnected = $false
        $Global:Tenant = ""
        
        $Window.Dispatcher.Invoke([Action]{
            $lblConnectionStatus.Text = "Disconnect Error"
            $lblConnectionStatus.Foreground = "Red"
        })
        
        Set-ControlsEnabled $false
        Update-Status "Disconnect Error"
    }
}

# Connect to Microsoft Graph
function Connect-Graph {
    try {
        if (-not $Global:GraphConnected) {
            Update-Log "Connecting to Microsoft Graph..."
            Connect-MgGraph -Scopes "User.ReadWrite.All" -ErrorAction Stop
            $Global:GraphConnected = $true
            Update-Log "Successfully connected to Microsoft Graph" "Green"
        }
        return $true
    } catch {
        Update-Log "Failed to connect to Microsoft Graph: $($_.Exception.Message)" "Red"
        return $false
    }
}

# Add SendAs permission
function Add-SendAsPermission {
    $Mailbox = $txtMailbox.Text.Trim()
    $Trustee = $txtTrustee.Text.Trim()
    
    if ([string]::IsNullOrEmpty($Mailbox) -or [string]::IsNullOrEmpty($Trustee)) {
        Update-Log "Please enter both mailbox and trustee." "Red"
        return
    }
    
    if (-not (Test-MailboxExists $Mailbox)) { return }
    
    try {
        Start-ActivityBar "Adding SendAs permission..."
        Add-RecipientPermission -Identity $Mailbox -Trustee $Trustee -AccessRights SendAs -Confirm:$false
        Stop-ActivityBar
        Update-Log "SendAs permission added for $Trustee on $Mailbox" "Green"
    } catch {
        Stop-ActivityBar
        Update-Log "Failed to add SendAs permission: $($_.Exception.Message)" "Red"
    }
}

# Remove SendAs permission
function Remove-SendAsPermission {
    $Mailbox = $txtMailbox.Text.Trim()
    $Trustee = $txtTrustee.Text.Trim()
    
    if ([string]::IsNullOrEmpty($Mailbox) -or [string]::IsNullOrEmpty($Trustee)) {
        Update-Log "Please enter both mailbox and trustee." "Red"
        return
    }
    
    if (-not (Test-MailboxExists $Mailbox)) { return }
    
    try {
        Start-ActivityBar "Removing SendAs permission..."
        Remove-RecipientPermission -Identity $Mailbox -Trustee $Trustee -AccessRights SendAs -Confirm:$false
        Stop-ActivityBar
        Update-Log "SendAs permission removed for $Trustee on $Mailbox" "Green"
    } catch {
        Stop-ActivityBar
        Update-Log "Failed to remove SendAs permission: $($_.Exception.Message)" "Red"
    }
}

# Add FullAccess permission
function Add-FullAccessPermission {
    $Mailbox = $txtMailbox.Text.Trim()
    $Trustee = $txtTrustee.Text.Trim()
    
    if ([string]::IsNullOrEmpty($Mailbox) -or [string]::IsNullOrEmpty($Trustee)) {
        Update-Log "Please enter both mailbox and trustee." "Red"
        return
    }
    
    if (-not (Test-MailboxExists $Mailbox)) { return }
    
    try {
        Start-ActivityBar "Adding FullAccess permission..."
        Add-MailboxPermission -Identity $Mailbox -User $Trustee -AccessRights FullAccess -AutoMapping:$false -Confirm:$false
        Stop-ActivityBar
        Update-Log "FullAccess permission added for $Trustee on $Mailbox" "Green"
    } catch {
        Stop-ActivityBar
        Update-Log "Failed to add FullAccess permission: $($_.Exception.Message)" "Red"
    }
}

# Remove FullAccess permission
function Remove-FullAccessPermission {
    $Mailbox = $txtMailbox.Text.Trim()
    $Trustee = $txtTrustee.Text.Trim()
    
    if ([string]::IsNullOrEmpty($Mailbox) -or [string]::IsNullOrEmpty($Trustee)) {
        Update-Log "Please enter both mailbox and trustee." "Red"
        return
    }
    
    if (-not (Test-MailboxExists $Mailbox)) { return }
    
    try {
        Start-ActivityBar "Removing FullAccess permission..."
        Remove-MailboxPermission -Identity $Mailbox -User $Trustee -AccessRights FullAccess -Confirm:$false
        Stop-ActivityBar
        Update-Log "FullAccess permission removed for $Trustee on $Mailbox" "Green"
    } catch {
        Stop-ActivityBar
        Update-Log "Failed to remove FullAccess permission: $($_.Exception.Message)" "Red"
    }
}

# Add SendOnBehalf permission
function Add-SendOnBehalfPermission {
    $Mailbox = $txtMailbox.Text.Trim()
    $Trustee = $txtTrustee.Text.Trim()
    
    if ([string]::IsNullOrEmpty($Mailbox) -or [string]::IsNullOrEmpty($Trustee)) {
        Update-Log "Please enter both mailbox and trustee." "Red"
        return
    }
    
    if (-not (Test-MailboxExists $Mailbox)) { return }
    
    try {
        Start-ActivityBar "Adding SendOnBehalf permission..."
        Set-Mailbox -Identity $Mailbox -GrantSendOnBehalfTo @{Add=$Trustee}
        Stop-ActivityBar
        Update-Log "SendOnBehalf permission added for $Trustee on $Mailbox" "Green"
    } catch {
        Stop-ActivityBar
        Update-Log "Failed to add SendOnBehalf permission: $($_.Exception.Message)" "Red"
    }
}

# Remove SendOnBehalf permission
function Remove-SendOnBehalfPermission {
    $Mailbox = $txtMailbox.Text.Trim()
    $Trustee = $txtTrustee.Text.Trim()
    
    if ([string]::IsNullOrEmpty($Mailbox) -or [string]::IsNullOrEmpty($Trustee)) {
        Update-Log "Please enter both mailbox and trustee." "Red"
        return
    }
    
    if (-not (Test-MailboxExists $Mailbox)) { return }
    
    try {
        Start-ActivityBar "Removing SendOnBehalf permission..."
        Set-Mailbox -Identity $Mailbox -GrantSendOnBehalfTo @{Remove=$Trustee}
        Stop-ActivityBar
        Update-Log "SendOnBehalf permission removed for $Trustee on $Mailbox" "Green"
    } catch {
        Stop-ActivityBar
        Update-Log "Failed to remove SendOnBehalf permission: $($_.Exception.Message)" "Red"
    }
}

# Add Calendar permission
function Add-CalendarPermission {
    $Mailbox = $txtMailbox.Text.Trim()
    $Trustee = $txtCalendarTrustee.Text.Trim()
    $Role = $cmbCalendarRole.SelectedItem.Content
    
    if ([string]::IsNullOrEmpty($Mailbox) -or [string]::IsNullOrEmpty($Trustee)) {
        Update-Log "Please enter both mailbox and trustee for calendar permission." "Red"
        return
    }
    
    if (-not (Test-MailboxExists $Mailbox)) { return }
    
    try {
        Start-ActivityBar "Adding Calendar permission..."
        Add-MailboxFolderPermission -Identity "${Mailbox}:\Calendar" -User $Trustee -AccessRights $Role -Confirm:$false
        Stop-ActivityBar
        Update-Log "Calendar permission ($Role) added for $Trustee on $Mailbox" "Green"
    } catch {
        Stop-ActivityBar
        Update-Log "Failed to add Calendar permission: $($_.Exception.Message)" "Red"
    }
}

# Remove Calendar permission
function Remove-CalendarPermission {
    $Mailbox = $txtMailbox.Text.Trim()
    $Trustee = $txtCalendarTrustee.Text.Trim()
    
    if ([string]::IsNullOrEmpty($Mailbox) -or [string]::IsNullOrEmpty($Trustee)) {
        Update-Log "Please enter both mailbox and trustee for calendar permission." "Red"
        return
    }
    
    if (-not (Test-MailboxExists $Mailbox)) { return }
    
    try {
        Start-ActivityBar "Removing Calendar permission..."
        Remove-MailboxFolderPermission -Identity "${Mailbox}:\Calendar" -User $Trustee -Confirm:$false
        Stop-ActivityBar
        Update-Log "Calendar permission removed for $Trustee on $Mailbox" "Green"
    } catch {
        Stop-ActivityBar
        Update-Log "Failed to remove Calendar permission: $($_.Exception.Message)" "Red"
    }
}

# List Calendar permissions
function Show-CalendarPermissions {
    $Mailbox = $txtMailbox.Text.Trim()
    
    if ([string]::IsNullOrEmpty($Mailbox)) {
        Update-Log "Please enter a mailbox." "Red"
        return
    }
    
    if (-not (Test-MailboxExists $Mailbox)) { return }
    
    try {
        Start-ActivityBar "Retrieving Calendar permissions..."
        $CalendarPermissions = Get-MailboxFolderPermission -Identity "${Mailbox}:\Calendar"
        Stop-ActivityBar
        Update-Log "Calendar permissions for $Mailbox`:" "Blue"
        foreach ($perm in $CalendarPermissions) {
            Update-Log "  User: $($perm.User) - Access Rights: $($perm.AccessRights -join ', ')" "Blue"
        }
    } catch {
        Stop-ActivityBar
        Update-Log "Failed to retrieve Calendar permissions: $($_.Exception.Message)" "Red"
    }
}

# List all permissions
function Show-AllPermissions {
    $Mailbox = $txtMailbox.Text.Trim()
    
    if ([string]::IsNullOrEmpty($Mailbox)) {
        Update-Log "Please enter a mailbox." "Red"
        return
    }
    
    if (-not (Test-MailboxExists $Mailbox)) { return }
    
    try {
        Start-ActivityBar "Retrieving all permissions..."
        Update-Log "=== SendAs Permissions ===" "Blue"
        $SendAsPerms = Get-RecipientPermission -Identity $Mailbox
        foreach ($perm in $SendAsPerms) {
            Update-Log "  Trustee: $($perm.Trustee) - Access Rights: $($perm.AccessRights -join ', ')" "Blue"
        }
        
        Update-Log "`n=== FullAccess Permissions ===" "Blue"
        $FullAccessPerms = Get-MailboxPermission -Identity $Mailbox | Where-Object { $_.AccessRights -contains "FullAccess" }
        foreach ($perm in $FullAccessPerms) {
            Update-Log "  User: $($perm.User) - Access Rights: $($perm.AccessRights -join ', ') - Deny: $($perm.Deny) - Inherited: $($perm.IsInherited)" "Blue"
        }
        
        Update-Log "`n=== SendOnBehalf Permissions ===" "Blue"
        $SendOnBehalfPerms = (Get-Mailbox -Identity $Mailbox).GrantSendOnBehalfTo
        foreach ($perm in $SendOnBehalfPerms) {
            Update-Log "  User: $perm" "Blue"
        }
        
        Update-Log "`n=== Calendar Permissions ===" "Blue"
        $CalendarPerms = Get-MailboxFolderPermission -Identity "${Mailbox}:\Calendar"
        foreach ($perm in $CalendarPerms) {
            Update-Log "  User: $($perm.User) - Access Rights: $($perm.AccessRights -join ', ')" "Blue"
        }
        Stop-ActivityBar
    } catch {
        Stop-ActivityBar
        Update-Log "Failed to retrieve permissions: $($_.Exception.Message)" "Red"
    }
}

# Validate mailbox
function Test-MailboxValidation {
    $Mailbox = $txtMailbox.Text.Trim()
    
    if ([string]::IsNullOrEmpty($Mailbox)) {
        Update-Log "Please enter a mailbox." "Red"
        return
    }
    
    try {
        Start-ActivityBar "Validating mailbox..."
        if (Test-MailboxExists $Mailbox) {
            Stop-ActivityBar
            Update-Log "Mailbox [$Mailbox] exists and is accessible." "Green"
        } else {
            Stop-ActivityBar
            Update-Log "Mailbox [$Mailbox] does not exist or is not accessible." "Red"
        }
    } catch {
        Stop-ActivityBar
        Update-Log "Error validating mailbox: $($_.Exception.Message)" "Red"
    }
}

# Initialize form state - disable all buttons except Connect
Set-ControlsEnabled $false

# Ensure Connect button is enabled and others are disabled
$Window.Dispatcher.Invoke([Action]{
    $btnConnect.IsEnabled = $true
    #Update-Log "Form initialized - all buttons disabled except Connect" "Blue"
})

# Event handlers
$btnConnect.Add_Click({
    Update-Log "Starting Connect to Exchange Online operation..." "Blue"
    Connect-Exchange
})

if ($btnRefreshTenant) {
    $btnRefreshTenant.Add_Click({
        Update-Log "Starting Refresh Tenant Info operation..." "Blue"
        Refresh-TenantInfo
    })
}


# Add click handler to clear mailbox field
$txtMailbox.Add_GotFocus({
    if ($txtMailbox.Text -eq "Enter target UPN here") {
        $txtMailbox.Text = ""
        #Update-Log "Cleared mailbox field for new entry" "Blue"
    }
})

# Add lost focus handler to restore placeholder text for mailbox field
$txtMailbox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($txtMailbox.Text)) {
        $txtMailbox.Text = "Enter target UPN here"
    }
})

# Add click handler to clear trustee field
$txtTrustee.Add_GotFocus({
    if ($txtTrustee.Text -eq "user@tenantname.com") {
        $txtTrustee.Text = ""
        #Update-Log "Cleared trustee field for new entry" "Blue"
    }
})

# Add click handler to clear trustee field
$txtCalendarTrustee.Add_GotFocus({
    if ($txtCalendarTrustee.Text -eq "Enter target UPN here") {
        $txtCalendarTrustee.Text = ""
        #Update-Log "Cleared trustee field for new entry" "Blue"
    }
})

# Add lost focus handler to restore placeholder text for calendar trustee field
$txtCalendarTrustee.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($txtCalendarTrustee.Text)) {
        $txtCalendarTrustee.Text = "Enter target UPN here"
    }
})

# Timer for auto-copying mailbox value to calendar trustee field
$Global:MailboxCopyTimer = $null

# Add text changed handler to mailbox field for auto-copy functionality
$txtMailbox.Add_TextChanged({
    # Clear existing timer if it exists
    if ($Global:MailboxCopyTimer) {
        $Global:MailboxCopyTimer.Stop()
    }
    
    # Create new timer with 1 second delay
    $Global:MailboxCopyTimer = New-Object System.Windows.Threading.DispatcherTimer
    $Global:MailboxCopyTimer.Interval = [TimeSpan]::FromSeconds(1)
    $Global:MailboxCopyTimer.Add_Tick({
        # Stop the timer
        $Global:MailboxCopyTimer.Stop()
        $Global:MailboxCopyTimer = $null
        
        # Copy mailbox value to calendar trustee field if it's not empty and not placeholder text
        $MailboxValue = $txtMailbox.Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($MailboxValue) -and $MailboxValue -ne "Enter target UPN here") {
            $txtCalendarTrustee.Text = $MailboxValue
            #Update-Log "Auto-copied mailbox value to calendar trustee field: $MailboxValue" "Blue"
        }
    })
    
    # Start the timer
    $Global:MailboxCopyTimer.Start()
})

$btnDisconnect.Add_Click({
    Update-Log "Starting Disconnect from Exchange Online operation..." "Blue"
    Disconnect-Exchange
})

$btnAddSendAs.Add_Click({
    Update-Log "Starting Add SendAs permission operation..." "Blue"
    Add-SendAsPermission
})

$btnRemoveSendAs.Add_Click({
    Update-Log "Starting Remove SendAs permission operation..." "Blue"
    Remove-SendAsPermission
})

$btnAddFullAccess.Add_Click({
    Update-Log "Starting Add FullAccess permission operation..." "Blue"
    Add-FullAccessPermission
})

$btnRemoveFullAccess.Add_Click({
    Update-Log "Starting Remove FullAccess permission operation..." "Blue"
    Remove-FullAccessPermission
})

$btnAddSendOnBehalf.Add_Click({
    Update-Log "Starting Add SendOnBehalf permission operation..." "Blue"
    Add-SendOnBehalfPermission
})

$btnRemoveSendOnBehalf.Add_Click({
    Update-Log "Starting Remove SendOnBehalf permission operation..." "Blue"
    Remove-SendOnBehalfPermission
})

$btnAddCalendar.Add_Click({
    Update-Log "Starting Add Calendar permission operation..." "Blue"
    Add-CalendarPermission
})

$btnRemoveCalendar.Add_Click({
    Update-Log "Starting Remove Calendar permission operation..." "Blue"
    Remove-CalendarPermission
})

$btnListCalendar.Add_Click({
    Update-Log "Starting List Calendar permissions operation..." "Blue"
    Show-CalendarPermissions
})

$btnListPermissions.Add_Click({
    Update-Log "Starting List All Permissions operation..." "Blue"
    Show-AllPermissions
})

$btnValidateMailbox.Add_Click({
    Update-Log "Starting Validate Mailbox operation..." "Blue"
    Test-MailboxValidation
})

$btnClearLog.Add_Click({
    Update-Log "Starting Clear Log operation..." "Blue"
    Start-ActivityBar "Clearing log..."
    $txtLog.Text = "Log cleared at $(Get-Date -Format 'HH:mm:ss')`n"
    Stop-ActivityBar
})

# Show the window
$Window.ShowDialog() | Out-Null

# Cleanup on close
Update-Log "Application closing - performing comprehensive cleanup..." "Blue"

# Disconnect from all services
if ($Global:ExchangeConnected -or $Global:GraphConnected) {
    try {
        # Use the enhanced disconnect function
        Disconnect-Exchange
    } catch {
        # Force cleanup even if disconnect fails
        try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        } catch { }
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
        } catch { }
        
        # Clear all sessions
        try {
            Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -or $_.ConfigurationName -like "*Graph*" } | Remove-PSSession -ErrorAction SilentlyContinue
        } catch { }
        
        Update-Log "Forced cleanup completed" "Yellow"
    }
}

# Cleanup timer
if ($Global:MailboxCopyTimer) {
    try {
        $Global:MailboxCopyTimer.Stop()
        $Global:MailboxCopyTimer = $null
    } catch {
        # Ignore timer cleanup errors
    }
}
