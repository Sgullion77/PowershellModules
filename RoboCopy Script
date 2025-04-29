# Get Date and Time
Get-Date

# Test if script is being run as admin, if not it will re-launch as admin
function Test-IsAdmin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    $adminRole = [Security.Principal.WindowsBuiltInRole]::Administrator
    return $currentPrincipal.IsInRole($adminRole)
}

if (-not (Test-IsAdmin)) {
    if ($env:USERNAME -ne "SYSTEM") {
        Write-Warning "This script needs to be run as an administrator."
        Pause
        Start-Process powershell.exe -Verb RunAs -ArgumentList "-File `"$($MyInvocation.MyCommand.Path)`"", $MyInvocation.UnboundArguments
        exit
    }
}

# Load the necessary WPF assemblies explicitly
Add-Type -AssemblyName "PresentationCore"
Add-Type -AssemblyName "PresentationFramework"

# Create the window
$window = New-Object Windows.Window
$window.Title = "Backup or Restore User Profile"
$window.Width = 400
$window.Height = 200
$window.WindowStartupLocation = 'CenterScreen'

# Create a StackPanel to organize elements
$stackPanel = New-Object Windows.Controls.StackPanel
$window.Content = $stackPanel

# Create the label
$label = New-Object Windows.Controls.Label
$label.Content = "Please select 'Backup' or 'Restore' from the dropdown list:"
$label.Margin = '10,10,10,10'
$stackPanel.Children.Add($label)

# Create the ComboBox (dropdown list)
$comboBox = New-Object Windows.Controls.ComboBox
$comboBox.Margin = '10,10,10,10'
$comboBox.Width = 300
$comboBox.Items.Add("Backup")
$comboBox.Items.Add("Restore")
$comboBox.SelectedIndex = 0  # Default to "Backup"
$stackPanel.Children.Add($comboBox)

# Create the Submit button
$button = New-Object Windows.Controls.Button
$button.Content = "Submit"
$button.Margin = '10,10,10,10'
$stackPanel.Children.Add($button)

# Add Button Click Event to process input
$button.Add_Click({
    # Get selected item from ComboBox
    $Action = $comboBox.SelectedItem

    # Close the window after input
    $window.Close()

    if ($Action -notin @('Backup', 'Restore')) {
        Write-Host "Invalid selection. Please choose 'Backup' or 'Restore'."
        exit
    }

    # Enter username
    $Username = Read-Host -Prompt "Please enter the username for the user you need to backup or restore"

    # Run robocopy commands based on selection
    switch ($Action.ToLower()) {
        'backup' {
            Write-Host "Starting BACKUP for $Username..."

            robocopy "C:\Users\$Username\AppData\Local\Google\Chrome\User Data\Default\Bookmarks.bak" "C:\$Username\Google Bookmarks" /E /COPYALL /R:3 /W:5
            robocopy "C:\Users\$Username\AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks" "C:\$Username\Edge Bookmarks" /E /COPYALL /R:3 /W:5
            robocopy "C:\Users\$Username\Desktop" "C:\$Username\Desktop" /E /COPYALL /R:3 /W:5
            robocopy "C:\Users\$Username\Documents" "C:\$Username\Documents" /E /COPYALL /R:3 /W:5
            robocopy "C:\Users\$Username\Downloads" "C:\$Username\Downloads" /E /COPYALL /R:3 /W:5
            robocopy "C:\Users\$Username\Pictures" "C:\$Username\Pictures" /E /COPYALL /R:3 /W:5
            robocopy "C:\Users\$Username\Favorites" "C:\$Username\Favorites" /E /COPYALL /R:3 /W:5
            robocopy "C:\Users\$Username\Videos" "C:\$Username\Videos" /E /COPYALL /R:3 /W:5
            robocopy "C:\Users\$Username\AppData\Roaming\Microsoft\Signatures" "C:\$Username\Signatures" /E /COPYALL /R:3 /W:5

            Write-Host "Backup completed."
			
			    # SID and registry rename section (unchanged)
    $user = New-Object System.Security.Principal.NTAccount($Username)
    try {
        $sid = $user.Translate([System.Security.Principal.SecurityIdentifier]).Value
    } catch {
        Write-Error "User '$Username' not found."
        exit
    }

    $profilePath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid"
    if (Test-Path $profilePath) {
        $newProfilePath = "$profilePath.old"
        Rename-Item -Path $profilePath -NewName "$sid.old"
        Write-Host "Renamed registry key '$profilePath' to '$newProfilePath'"
    } else {
        Write-Warning "Registry key '$profilePath' not found."
    }
        }

        'restore' {
            Write-Host "Starting RESTORE for $Username..."

            robocopy "C:\$Username\Google Bookmarks" "C:\Users\$Username\AppData\Local\Google\Chrome\User Data\Default" /E /COPYALL /R:3 /W:5
            robocopy "C:\$Username\Edge Bookmarks" "C:\Users\$Username\AppData\Local\Microsoft\Edge\User Data\Default" /E /COPYALL /R:3 /W:5
            robocopy "C:\$Username\Desktop" "C:\Users\$Username\Desktop" /E /COPYALL /R:3 /W:5
            robocopy "C:\$Username\Documents" "C:\Users\$Username\Documents" /E /COPYALL /R:3 /W:5
            robocopy "C:\$Username\Downloads" "C:\Users\$Username\Downloads" /E /COPYALL /R:3 /W:5
            robocopy "C:\$Username\Pictures" "C:\Users\$Username\Pictures" /E /COPYALL /R:3 /W:5
            robocopy "C:\$Username\Favorites" "C:\Users\$Username\Favorites" /E /COPYALL /R:3 /W:5
            robocopy "C:\$Username\Videos" "C:\Users\$Username\Videos" /E /COPYALL /R:3 /W:5
            robocopy "C:\$Username\Signatures" "C:\Users\$Username\AppData\Roaming\Microsoft\Signatures" /E /COPYALL /R:3 /W:5

            Write-Host "Restore completed."
        }
    }



    # Ask for reboot
    Do {
        $InputReboot = Read-Host "Do you wish to reboot now? [y/n]"
    } While ($InputReboot -notin @('y', 'n'))

    Switch ($InputReboot.ToLower()) {
        "y" {Restart-Computer -Force}
        "n" {exit}
    }
})

# Show the window
$window.ShowDialog()
