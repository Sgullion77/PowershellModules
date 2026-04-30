# Load assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
if (-not (Test-IsAdmin)) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-File `"$($MyInvocation.MyCommand.Path)`""
    exit
}

function Select-Folder {
    $f = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($f.ShowDialog() -eq "OK") { return $f.SelectedPath }
    return $null
}

function Write-Log {
    param($msg)
    $time = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    "$time - $msg" | Out-File -Append -FilePath $logFile
}

function Close-UserApps {
    Write-Host "Closing apps in 5 seconds..."
    Start-Sleep 5

    $processes = @("chrome","msedge","firefox","outlook","teams","onedrive","explorer")

    foreach ($p in $processes) {
        Get-Process -Name $p -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                if ($_.CloseMainWindow()) { Start-Sleep 2 }
                if (!$_.HasExited) { Stop-Process $_ -Force }
                Write-Log "Closed $($_.ProcessName)"
            } catch {}
        }
    }

    Start-Sleep 2
    Start-Process explorer.exe
    Write-Log "Explorer restarted"
}

function Run-Robo {
    param($src,$dst,$name)

    if (Test-Path $src) {
        if (!(Test-Path $dst)) {
            New-Item -ItemType Directory -Path $dst -Force | Out-Null
        }

        Write-Log "Copying $name"
        robocopy $src $dst /E /COPY:DAT /R:2 /W:2 /MT:16
    } else {
        Write-Log "Skipped $name (not found)"
    }
}

# Get users
$users = Get-ChildItem C:\Users | Where-Object {
    $_.Name -notin @("Public","Default","Default User","All Users")
} | Select-Object -ExpandProperty Name

# UI
$window = New-Object Windows.Window
$window.Title = "MSP Migration Tool"
$window.Width = 500
$window.Height = 520

$stack = New-Object Windows.Controls.StackPanel
$window.Content = $stack

$userBox = New-Object Windows.Controls.ComboBox
$users | ForEach-Object { $userBox.Items.Add($_) }
$userBox.SelectedIndex = 0
$userBox.Margin = "10"
$stack.Children.Add($userBox)

$actionBox = New-Object Windows.Controls.ComboBox
$actionBox.Items.Add("Backup")
$actionBox.Items.Add("Restore")
$actionBox.SelectedIndex = 0
$actionBox.Margin = "10"
$stack.Children.Add($actionBox)

$checks = @{}
function Add-Check($name) {
    $cb = New-Object Windows.Controls.CheckBox
    $cb.Content = $name
    $cb.IsChecked = $true
    $cb.Margin = "5"
    $stack.Children.Add($cb)
    $checks[$name] = $cb
}

"Desktop","Documents","Pictures","Videos","Favorites","Downloads",
"Chrome","Edge","Firefox","Outlook","Signatures","Start Menu",
"Quick Access","Taskbar Pins" | ForEach-Object { Add-Check $_ }

$progress = New-Object Windows.Controls.ProgressBar
$progress.Height = 20
$progress.Margin = "10"
$progress.Minimum = 0
$progress.Maximum = 100
$stack.Children.Add($progress)

$status = New-Object Windows.Controls.Label
$status.Content = "Ready"
$stack.Children.Add($status)

$btn = New-Object Windows.Controls.Button
$btn.Content = "Run"
$btn.Margin = "10"
$stack.Children.Add($btn)

$btn.Add_Click({

    $user = $userBox.SelectedItem
    $action = $actionBox.SelectedItem

    $base = Select-Folder
    if (-not $base) { return }

    $global:logFile = "$base\migration_log.txt"
    New-Item $logFile -Force | Out-Null

    Close-UserApps

    # FIXED TASK LIST
    $tasks = New-Object System.Collections.ArrayList

    function Add-Task($name,$src,$dst) {
        if ($checks[$name].IsChecked -eq $true) {
            $null = $tasks.Add(@{n=$name;s=$src;d=$dst})
        }
    }

    if ($action -eq "Backup") {
        Add-Task "Desktop" "C:\Users\$user\Desktop" "$base\Desktop"
        Add-Task "Documents" "C:\Users\$user\Documents" "$base\Documents"
        Add-Task "Pictures" "C:\Users\$user\Pictures" "$base\Pictures"
        Add-Task "Videos" "C:\Users\$user\Videos" "$base\Videos"
        Add-Task "Favorites" "C:\Users\$user\Favorites" "$base\Favorites"
        Add-Task "Downloads" "C:\Users\$user\Downloads" "$base\Downloads"

        Add-Task "Chrome" "C:\Users\$user\AppData\Local\Google\Chrome\User Data" "$base\Chrome"
        Add-Task "Edge" "C:\Users\$user\AppData\Local\Microsoft\Edge\User Data" "$base\Edge"
        Add-Task "Firefox" "C:\Users\$user\AppData\Roaming\Mozilla\Firefox\Profiles" "$base\Firefox"

        Add-Task "Outlook" "C:\Users\$user\AppData\Local\Microsoft\Outlook" "$base\Outlook"
        Add-Task "Signatures" "C:\Users\$user\AppData\Roaming\Microsoft\Signatures" "$base\Signatures"
        Add-Task "Start Menu" "C:\Users\$user\AppData\Roaming\Microsoft\Windows\Start Menu" "$base\StartMenu"

        Add-Task "Quick Access" "C:\Users\$user\AppData\Roaming\Microsoft\Windows\Recent" "$base\QuickAccess"
        Add-Task "Taskbar Pins" "C:\Users\$user\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar" "$base\Taskbar"
    }
    else {
        Add-Task "Desktop" "$base\Desktop" "C:\Users\$user\Desktop"
        Add-Task "Documents" "$base\Documents" "C:\Users\$user\Documents"
        Add-Task "Pictures" "$base\Pictures" "C:\Users\$user\Pictures"
        Add-Task "Videos" "$base\Videos" "C:\Users\$user\Videos"
        Add-Task "Favorites" "$base\Favorites" "C:\Users\$user\Favorites"
        Add-Task "Downloads" "$base\Downloads" "C:\Users\$user\Downloads"

        Add-Task "Chrome" "$base\Chrome" "C:\Users\$user\AppData\Local\Google\Chrome\User Data"
        Add-Task "Edge" "$base\Edge" "C:\Users\$user\AppData\Local\Microsoft\Edge\User Data"
        Add-Task "Firefox" "$base\Firefox" "C:\Users\$user\AppData\Roaming\Mozilla\Firefox\Profiles"

        Add-Task "Outlook" "$base\Outlook" "C:\Users\$user\AppData\Local\Microsoft\Outlook"
        Add-Task "Signatures" "$base\Signatures" "C:\Users\$user\AppData\Roaming\Microsoft\Signatures"
        Add-Task "Start Menu" "$base\StartMenu" "C:\Users\$user\AppData\Roaming\Microsoft\Windows\Start Menu"

        Add-Task "Quick Access" "$base\QuickAccess" "C:\Users\$user\AppData\Roaming\Microsoft\Windows\Recent"
        Add-Task "Taskbar Pins" "$base\Taskbar" "C:\Users\$user\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
    }

    $total = $tasks.Count
    $i = 0

    if ($total -eq 0) {
        $status.Content = "Nothing selected"
        return
    }

    foreach ($t in $tasks) {
        $i++

        $status.Content = "Processing $($t.n)..."
        Run-Robo $t.s $t.d $t.n

        $progress.Value = [math]::Round(($i / $total) * 100)

        $window.Dispatcher.Invoke([System.Action]{}, "Render")
    }

    $status.Content = "Completed"
    Write-Log "DONE"
})

$window.ShowDialog()
