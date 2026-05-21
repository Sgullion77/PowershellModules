# ============================================
# FlexTG Profile Migration Tool
# ============================================

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Global logfile variable
$script:logFile = $null

# ============================================
# Admin Check
# ============================================

function Test-IsAdmin {

    $id = [Security.Principal.WindowsIdentity]::GetCurrent()

    $p = New-Object Security.Principal.WindowsPrincipal($id)

    return $p.IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator
    )
}

if (-not (Test-IsAdmin)) {

    Start-Process powershell.exe `
        -Verb RunAs `
        -ArgumentList "-ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`""

    exit
}

# ============================================
# Logging
# ============================================

function Write-Log {

    param([string]$Message)

    $time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Prevent null logfile crash
    if ([string]::IsNullOrWhiteSpace($script:logFile)) {

        Write-Host "$time - $Message"

        return
    }

    try {

        "$time - $Message" | Out-File `
            -FilePath $script:logFile `
            -Append `
            -Encoding utf8

    } catch {

        Write-Host "LOG ERROR: $_"
    }
}

# ============================================
# Folder Picker
# ============================================

function Select-Folder {

    $f = New-Object System.Windows.Forms.FolderBrowserDialog

    if ($f.ShowDialog() -eq "OK") {
        return $f.SelectedPath
    }

    return $null
}

# ============================================
# Close Applications
# ============================================

function Close-UserApps {

    Write-Host "Closing user applications in 5 seconds..."
    Start-Sleep 5

    Write-Log "Closing user applications"

    $processes = @(
        "chrome",
        "msedge",
        "firefox",
        "outlook",
        "teams",
        "onedrive",
        "explorer"
    )

    foreach ($procName in $processes) {

        $procs = Get-Process `
            -Name $procName `
            -ErrorAction SilentlyContinue

        foreach ($proc in $procs) {

            try {

                if ($proc.CloseMainWindow()) {

                    Write-Log "Gracefully closed $($proc.ProcessName)"

                    Start-Sleep 2
                }

                if (-not $proc.HasExited) {

                    Stop-Process `
                        -Id $proc.Id `
                        -Force

                    Write-Log "Force killed $($proc.ProcessName)"
                }

            } catch {

                Write-Log "Failed closing $($proc.ProcessName): $_"
            }
        }
    }

    Start-Sleep 2

    Start-Process explorer.exe

    Write-Log "Explorer restarted"
}

# ============================================
# Robocopy Wrapper
# ============================================

function Run-Robo {

    param(
        [string]$Source,
        [string]$Destination,
        [string]$Name
    )

    try {

        if (-not (Test-Path $Source)) {

            Write-Log "Skipped $Name (source missing)"

            return
        }

        if (-not (Test-Path $Destination)) {

            New-Item `
                -ItemType Directory `
                -Path $Destination `
                -Force | Out-Null
        }

        Write-Log "Copying $Name"

        robocopy `
            $Source `
            $Destination `
            /E `
            /COPY:DAT `
            /R:2 `
            /W:2 `
            /MT:16 `
            /NFL `
            /NDL `
            /NJH `
            /NJS `
            /NP | Out-Null

        Write-Log "Finished $Name"

    } catch {

        Write-Log "ERROR copying $Name : $_"
    }
}

# ============================================
# Get Users
# ============================================

$users = Get-ChildItem "C:\Users" | Where-Object {

    $_.PSIsContainer -and
    $_.Name -notin @(
        "Public",
        "Default",
        "Default User",
        "All Users"
    )

} | Select-Object -ExpandProperty Name

# ============================================
# WPF Window
# ============================================

$window = New-Object Windows.Window
$window.Title = "FlexTG Profile Migration Tool"
$window.Width = 500
$window.Height = 650
$window.WindowStartupLocation = "CenterScreen"

$stack = New-Object Windows.Controls.StackPanel

$window.Content = $stack

# ============================================
# User Dropdown
# ============================================

$userBox = New-Object Windows.Controls.ComboBox
$userBox.Margin = "10"

foreach ($u in $users) {

    [void]$userBox.Items.Add($u)
}

$userBox.SelectedIndex = 0

$stack.Children.Add($userBox)

# ============================================
# Action Dropdown
# ============================================

$actionBox = New-Object Windows.Controls.ComboBox
$actionBox.Margin = "10"

[void]$actionBox.Items.Add("Backup")
[void]$actionBox.Items.Add("Restore")

$actionBox.SelectedIndex = 0

$stack.Children.Add($actionBox)

# ============================================
# Checkboxes
# ============================================

$checks = @{}

function Add-Check {

    param([string]$Name)

    $cb = New-Object Windows.Controls.CheckBox

    $cb.Content = $Name
    $cb.IsChecked = $true
    $cb.Margin = "5"

    $stack.Children.Add($cb)

    $checks[$Name] = $cb
}

Add-Check "Desktop"
Add-Check "Documents"
Add-Check "Pictures"
Add-Check "Videos"
Add-Check "Favorites"
Add-Check "Downloads"

Add-Check "Chrome"
Add-Check "Edge"
Add-Check "Firefox"

Add-Check "Outlook"
Add-Check "Signatures"

Add-Check "Start Menu"
Add-Check "Quick Access"
Add-Check "Taskbar Pins"

# ============================================
# Progress Bar
# ============================================

$progress = New-Object Windows.Controls.ProgressBar

$progress.Height = 20
$progress.Margin = "10"
$progress.Minimum = 0
$progress.Maximum = 100

$stack.Children.Add($progress)

# ============================================
# Status Label
# ============================================

$status = New-Object Windows.Controls.Label

$status.Content = "Ready"
$status.Margin = "10"

$stack.Children.Add($status)

# ============================================
# Run Button
# ============================================

$btn = New-Object Windows.Controls.Button

$btn.Content = "Run"
$btn.Margin = "10"

$stack.Children.Add($btn)

# ============================================
# Run Logic
# ============================================

$btn.Add_Click({

    try {

        $user = $userBox.SelectedItem
        $action = $actionBox.SelectedItem

        $base = Select-Folder

        if (-not $base) {
            return
        }

        # Initialize logfile
        $script:logFile = Join-Path `
            $base `
            "migration_log.txt"

        New-Item `
            -ItemType File `
            -Path $script:logFile `
            -Force | Out-Null

        Write-Log "====================================="
        Write-Log "Migration started"
        Write-Log "User: $user"
        Write-Log "Action: $action"

        # Close applications
        Close-UserApps

        # ============================================
        # Build Task List
        # ============================================

        $script:tasks = @()

        function Add-Task {

            param(
                [string]$Name,
                [string]$Source,
                [string]$Destination
            )

            if ($checks[$Name].IsChecked -eq $true) {

                $script:tasks += [PSCustomObject]@{

                    Name        = $Name
                    Source      = $Source
                    Destination = $Destination
                }

                Write-Log "Queued task: $Name"
            }
        }

        # ============================================
        # Backup Tasks
        # ============================================

        if ($action -eq "Backup") {

            Add-Task "Desktop" "C:\Users\$user\Desktop" "$base\Desktop"
            Add-Task "Documents" "C:\Users\$user\Documents" "$base\Documents"
            Add-Task "Pictures" "C:\Users\$user\Pictures" "$base\Pictures"
            Add-Task "Videos" "C:\Users\$user\Videos" "$base\Videos"
            Add-Task "Favorites" "C:\Users\$user\Favorites" "$base\Favorites"
            Add-Task "Downloads" "C:\Users\$user\Downloads" "$base\Downloads"

            Add-Task "Chrome" `
                "C:\Users\$user\AppData\Local\Google\Chrome\User Data" `
                "$base\Chrome"

            Add-Task "Edge" `
                "C:\Users\$user\AppData\Local\Microsoft\Edge\User Data" `
                "$base\Edge"

            Add-Task "Firefox" `
                "C:\Users\$user\AppData\Roaming\Mozilla\Firefox\Profiles" `
                "$base\Firefox"

            Add-Task "Outlook" `
                "C:\Users\$user\AppData\Local\Microsoft\Outlook" `
                "$base\Outlook"

            Add-Task "Signatures" `
                "C:\Users\$user\AppData\Roaming\Microsoft\Signatures" `
                "$base\Signatures"

            Add-Task "Start Menu" `
                "C:\Users\$user\AppData\Roaming\Microsoft\Windows\Start Menu" `
                "$base\StartMenu"

            Add-Task "Quick Access" `
                "C:\Users\$user\AppData\Roaming\Microsoft\Windows\Recent" `
                "$base\QuickAccess"

            Add-Task "Taskbar Pins" `
                "C:\Users\$user\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar" `
                "$base\Taskbar"

        } else {

            # Restore Tasks

            Add-Task "Desktop" "$base\Desktop" "C:\Users\$user\Desktop"
            Add-Task "Documents" "$base\Documents" "C:\Users\$user\Documents"
            Add-Task "Pictures" "$base\Pictures" "C:\Users\$user\Pictures"
            Add-Task "Videos" "$base\Videos" "C:\Users\$user\Videos"
            Add-Task "Favorites" "$base\Favorites" "C:\Users\$user\Favorites"
            Add-Task "Downloads" "$base\Downloads" "C:\Users\$user\Downloads"

            Add-Task "Chrome" `
                "$base\Chrome" `
                "C:\Users\$user\AppData\Local\Google\Chrome\User Data"

            Add-Task "Edge" `
                "$base\Edge" `
                "C:\Users\$user\AppData\Local\Microsoft\Edge\User Data"

            Add-Task "Firefox" `
                "$base\Firefox" `
                "C:\Users\$user\AppData\Roaming\Mozilla\Firefox\Profiles"

            Add-Task "Outlook" `
                "$base\Outlook" `
                "C:\Users\$user\AppData\Local\Microsoft\Outlook"

            Add-Task "Signatures" `
                "$base\Signatures" `
                "C:\Users\$user\AppData\Roaming\Microsoft\Signatures"

            Add-Task "Start Menu" `
                "$base\StartMenu" `
                "C:\Users\$user\AppData\Roaming\Microsoft\Windows\Start Menu"

            Add-Task "Quick Access" `
                "$base\QuickAccess" `
                "C:\Users\$user\AppData\Roaming\Microsoft\Windows\Recent"

            Add-Task "Taskbar Pins" `
                "$base\Taskbar" `
                "C:\Users\$user\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
        }

        # ============================================
        # Execute Tasks
        # ============================================

        $total = $script:tasks.Count
        $current = 0

        Write-Log "Task count: $total"

        foreach ($task in $script:tasks) {

            $current++

            $status.Content = "Processing $($task.Name)..."

            Run-Robo `
                -Source $task.Source `
                -Destination $task.Destination `
                -Name $task.Name

            $progress.Value = (($current / $total) * 100)

            $window.Dispatcher.Invoke([Action]{}, "Render")
        }

        $status.Content = "Completed"
        $progress.Value = 100

        Write-Log "Migration completed"

        [System.Windows.MessageBox]::Show(
            "Migration completed successfully.",
            "Done",
            "OK",
            "Information"
        )

    } catch {

        Write-Log "FATAL ERROR: $_"

        [System.Windows.MessageBox]::Show(
            $_,
            "Error",
            "OK",
            "Error"
        )
    }
})

# ============================================
# Launch Window
# ============================================

$window.ShowDialog()
