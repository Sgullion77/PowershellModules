function Set-MailboxPermissions {
    # Prompt for tenant first
    $Tenant = Read-Host "Enter the Tenant Name (e.g. flextg.com, dekalbhousing.org)"

    # Initial connection to Exchange Online
    try {
        Connect-ExchangeOnline -Organization $Tenant -ErrorAction Stop
    } catch {
        Write-Host "Failed to connect to tenant $Tenant. Exiting..."
        return
    }

    # Setup logging
    $LogPath = "C:\Temp\Powershell-Logging"
    $LogFile = "$LogPath\MailboxPermissionLog.txt"
    if (!(Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath -Force | Out-Null }

    $CurrentUser = "$env:USERDOMAIN\$env:USERNAME"
    $GraphConnected = $false   # Track Graph connection state

    function Write-Log {
        param([string]$Message)
        $TimeStamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        Add-Content -Path $LogFile -Value "$TimeStamp - $CurrentUser - $Message"
    }

    # Prompt for initial mailbox + trustee
    $Identity = Read-Host "Enter the initial mailbox Identity (e.g. dispatch@dekalbhousing.org)"
    $Trustee = Read-Host "Enter the initial default Trustee (e.g. user@dekalbhousing.org)"

    do {
        Clear-Host
        Write-Host "================= Mailbox Permission Menu ============================        Tenant: $Tenant"
        Write-Host "Mailbox: $Identity        Trustee: $Trustee"
        Write-Host "======================================================================"
        Write-Host " 1.  Add SendAs Permission"
        Write-Host " 2.  Add FullAccess Permission"
        Write-Host " 3.  Add SendOnBehalf Permission"
        Write-Host " 4.  Remove SendAs Permission"
        Write-Host " 5.  Remove FullAccess Permission"
        Write-Host " 6.  Remove SendOnBehalf Permission"
        Write-Host " 7.  Exit"
        Write-Host " 8.  List Current Permissions"
        Write-Host " 9.  Add Calendar Permission (single trustee)"
        Write-Host "10. Remove Calendar Permission (single trustee)"
        Write-Host "11. List Calendar Permissions"
        Write-Host "12. Bulk Add Calendar Permissions (comma-separated trustees)"
        Write-Host "13. Bulk Add Calendar Permissions from CSV"
        Write-Host "14. Bulk Remove Calendar Permissions from CSV"
        Write-Host "15. Change Mailbox Identity (currently: $Identity)"
        Write-Host "16. Change Default Trustee (currently: $Trustee)"
        Write-Host "17. Change Tenant Name (and reconnect)"
        Write-Host "18. Change Primary UPN and Add Aliases (Graph API)"
        Write-Host "19. List Current UPN and Aliases (Graph API)"
        Write-Host "======================================================================"

        $choice = Read-Host "Select an option (1-19)"

        switch ($choice) {
            "1" {
                Add-RecipientPermission -Identity $Identity -Trustee $Trustee -AccessRights SendAs -Confirm:$false
                Write-Host "SendAs permission added."
                Write-Log "Added SendAs permission for Trustee [$Trustee] on [$Identity]"
            }
            "2" {
                Add-MailboxPermission -Identity $Identity -User $Trustee -AccessRights FullAccess -AutoMapping:$false -Confirm:$false
                Write-Host "FullAccess permission added."
                Write-Log "Added FullAccess permission for Trustee [$Trustee] on [$Identity]"
            }
            "3" {
                Set-Mailbox -Identity $Identity -GrantSendOnBehalfTo @{Add=$Trustee}
                Write-Host "SendOnBehalf permission added."
                Write-Log "Added SendOnBehalf permission for Trustee [$Trustee] on [$Identity]"
            }
            "4" {
                Remove-RecipientPermission -Identity $Identity -Trustee $Trustee -AccessRights SendAs -Confirm:$false
                Write-Host "SendAs permission removed."
                Write-Log "Removed SendAs permission for Trustee [$Trustee] on [$Identity]"
            }
            "5" {
                Remove-MailboxPermission -Identity $Identity -User $Trustee -AccessRights FullAccess -Confirm:$false
                Write-Host "FullAccess permission removed."
                Write-Log "Removed FullAccess permission for Trustee [$Trustee] on [$Identity]"
            }
            "6" {
                Set-Mailbox -Identity $Identity -GrantSendOnBehalfTo @{Remove=$Trustee}
                Write-Host "SendOnBehalf permission removed."
                Write-Log "Removed SendOnBehalf permission for Trustee [$Trustee] on [$Identity]"
            }
            "7" {
                Write-Host "Exiting..."
                break
            }
            "8" {
                Write-Host "`n--- SendAs Permissions ---"
                Get-RecipientPermission -Identity $Identity | Format-Table Trustee, AccessRights -AutoSize

                Write-Host "`n--- FullAccess Permissions ---"
                Get-MailboxPermission -Identity $Identity | Where-Object { $_.AccessRights -contains "FullAccess" } | Format-Table User, AccessRights, Deny, IsInherited -AutoSize

                Write-Host "`n--- SendOnBehalf Permissions ---"
                (Get-Mailbox -Identity $Identity).GrantSendOnBehalfTo | Format-Table -AutoSize

                Write-Host "`n--- Calendar Permissions ---"
                Get-MailboxFolderPermission -Identity "$Identity:\Calendar" | Format-Table User, AccessRights -AutoSize
            }
            "9" {
                $AccessRight = Read-Host "Enter Calendar Permission Role (e.g. Reviewer, Editor, Owner)"
                Add-MailboxFolderPermission -Identity "$Identity:\Calendar" -User $Trustee -AccessRights $AccessRight -Confirm:$false
                Write-Host "Calendar permission ($AccessRight) added for $Trustee."
                Write-Log "Added Calendar permission [$AccessRight] for Trustee [$Trustee] on [$Identity]"
            }
            "10" {
                Remove-MailboxFolderPermission -Identity "$Identity:\Calendar" -User $Trustee -Confirm:$false
                Write-Host "Calendar permission removed for $Trustee."
                Write-Log "Removed Calendar permission for Trustee [$Trustee] on [$Identity]"
            }
            "11" {
                Write-Host "Listing Calendar permissions for mailbox: $Identity"
                Get-MailboxFolderPermission -Identity "$Identity:\Calendar" | Format-Table User, AccessRights -AutoSize
            }
            "12" {
                $Trustees = Read-Host "Enter multiple trustees (comma-separated emails)"
                $TrusteeList = $Trustees -split ","
                $AccessRight = Read-Host "Enter Calendar Permission Role"
                foreach ($user in $TrusteeList) {
                    $trimmedUser = $user.Trim()
                    Add-MailboxFolderPermission -Identity "$Identity:\Calendar" -User $trimmedUser -AccessRights $AccessRight -Confirm:$false
                    Write-Host "Calendar permission ($AccessRight) added for $trimmedUser."
                    Write-Log "Added Calendar permission [$AccessRight] for Trustee [$trimmedUser] on [$Identity]"
                }
            }
            "13" {
                $CSVPath = Read-Host "Enter path to CSV (must have a 'User' column)"
                if (Test-Path $CSVPath) {
                    $Users = Import-Csv $CSVPath
                    $AccessRight = Read-Host "Enter Calendar Permission Role"
                    foreach ($u in $Users) {
                        $UserEmail = $u.User.Trim()
                        Add-MailboxFolderPermission -Identity "$Identity:\Calendar" -User $UserEmail -AccessRights $AccessRight -Confirm:$false
                        Write-Host "Calendar permission ($AccessRight) added for $UserEmail."
                        Write-Log "Added Calendar permission [$AccessRight] for Trustee [$UserEmail] on [$Identity]"
                    }
                } else {
                    Write-Host "CSV file not found at $CSVPath"
                }
            }
            "14" {
                $CSVPath = Read-Host "Enter path to CSV (must have a 'User' column)"
                if (Test-Path $CSVPath) {
                    $Users = Import-Csv $CSVPath
                    foreach ($u in $Users) {
                        $UserEmail = $u.User.Trim()
                        try {
                            Remove-MailboxFolderPermission -Identity "$Identity:\Calendar" -User $UserEmail -Confirm:$false -ErrorAction Stop
                            Write-Host "Calendar permission removed for $UserEmail."
                            Write-Log "Removed Calendar permission for Trustee [$UserEmail] on [$Identity]"
                        } catch {
                            Write-Host "Failed to remove permission for $UserEmail"
                            Write-Log "FAILED removing Calendar permission for [$UserEmail] on [$Identity]"
                        }
                    }
                } else {
                    Write-Host "CSV file not found at $CSVPath"
                }
            }
            "15" {
                $Identity = Read-Host "Enter NEW mailbox Identity"
                Write-Host "Mailbox identity changed to: $Identity"
                Write-Log "Changed mailbox identity to [$Identity]"
            }
            "16" {
                $Trustee = Read-Host "Enter NEW default Trustee"
                Write-Host "Trustee changed to: $Trustee"
                Write-Log "Changed default trustee to [$Trustee]"
            }
            "17" {
                $NewTenant = Read-Host "Enter NEW Tenant Name (e.g. dekalbhousing.org)"
                Write-Host "Disconnecting from current tenant..."
                Disconnect-ExchangeOnline -Confirm:$false
                try {
                    Connect-ExchangeOnline -Organization $NewTenant -ErrorAction Stop
                    $Tenant = $NewTenant
                    Write-Host "Reconnected to tenant: $Tenant"
                    Write-Log "Changed and reconnected to tenant [$Tenant]"

                    $Identity = Read-Host "Enter NEW mailbox Identity for tenant [$Tenant]"
                    $Trustee = Read-Host "Enter NEW default Trustee for tenant [$Tenant]"
                    Write-Log "Updated Identity to [$Identity] and Trustee to [$Trustee] for tenant [$Tenant]"

                } catch {
                    Write-Host "Failed to connect to tenant $NewTenant. Staying connected to previous tenant: $Tenant."
                    Write-Log "Failed to connect to new tenant [$NewTenant]"
                }
            }
            "18" {
                if (-not $GraphConnected) {
                    try {
                        Connect-MgGraph -Scopes "User.ReadWrite.All" -ErrorAction Stop
                        $GraphConnected = $true
                        Write-Host "Connected to Microsoft Graph."
                        Write-Log "Connected to Microsoft Graph for tenant [$Tenant]"
                    } catch {
                        Write-Host "Failed to connect to Graph API."
                        break
                    }
                }

                $NewPrimary = Read-Host "Enter NEW primary UPN for user [$Identity] (press Enter to skip)"
                $NewAliasesInput = Read-Host "Enter one or more NEW secondary aliases (comma-separated, press Enter to skip)"
                $AliasList = $NewAliasesInput -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

                try {
                    $User = Get-MgUser -UserId $Identity -Property OtherMails,UserPrincipalName
                    $CurrentAliases = $User.OtherMails

                    if ($AliasList.Count -gt 0) {
                        $UpdatedAliases = ($CurrentAliases + $AliasList) | Sort-Object -Unique
                    } else {
                        $UpdatedAliases = $CurrentAliases
                    }

                    $UpdateParams = @{}
                    if ($NewPrimary -and $NewPrimary -ne $User.UserPrincipalName) {
                        $UpdateParams["UserPrincipalName"] = $NewPrimary
                    }
                    if ($UpdatedAliases) {
                        $UpdateParams["OtherMails"] = $UpdatedAliases
                    }

                    if ($UpdateParams.Count -gt 0) {
                        Update-MgUser -UserId $Identity @UpdateParams

                        if ($UpdateParams.ContainsKey("UserPrincipalName")) {
                            Write-Host "Primary UPN changed from [$($User.UserPrincipalName)] to [$NewPrimary]"
                            Write-Log "Primary UPN changed from [$($User.UserPrincipalName)] to [$NewPrimary]"
                            $Identity = $NewPrimary
                        }
                        if ($UpdateParams.ContainsKey("OtherMails")) {
                            Write-Host "Aliases updated for [$Identity]: $($UpdatedAliases -join ', ')"
                            Write-Log "Updated aliases for [$Identity]: $($UpdatedAliases -join ', ')"
                        }
                    } else {
                        Write-Host "No changes made for [$Identity]."
                    }
                } catch {
                    Write-Host "Failed to update UPN/Aliases for [$Identity]"
                    Write-Log "FAILED updating UPN/Aliases for [$Identity]"
                }
            }
            "19" {
                if (-not $GraphConnected) {
                    try {
                        Connect-MgGraph -Scopes "User.Read.All" -ErrorAction Stop
                        $GraphConnected = $true
                        Write-Host "Connected to Microsoft Graph."
                        Write-Log "Connected to Microsoft Graph for tenant [$Tenant]"
                    } catch {
                        Write-Host "Failed to connect to Graph API."
                        break
                    }
                }

                try {
                    $User = Get-MgUser -UserId $Identity -Property OtherMails,UserPrincipalName
                    $PrimaryUPN = $User.UserPrincipalName
                    $SecondaryAliases = $User.OtherMails

                    $AliasTable = @()
                    $AliasTable += [PSCustomObject]@{Type="Primary"; Alias=$PrimaryUPN}
                    foreach ($a in $SecondaryAliases) {
                        $AliasTable += [PSCustomObject]@{Type="Secondary"; Alias=$a}
                    }
                    $AliasTable | Format-Table -AutoSize
                    Write-Log "Listed UPN/Aliases for [$Identity]"
                } catch {
                    Write-Host "Failed to retrieve UPN/Aliases for [$Identity]"
                    Write-Log "FAILED retrieving UPN/Aliases for [$Identity]"
                }
            }
            default {
                Write-Host "Invalid option. Try again."
            }
        }

        if ($choice -ne "7") {
            Pause
        }

    } while ($choice -ne "7")
}

# Run the function
Set-MailboxPermissions
