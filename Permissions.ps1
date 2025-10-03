function Set-MailboxPermissionsMenu {
    param()

    # Connect to Exchange Online
    Connect-ExchangeOnline

    # Setup logging
    $LogPath = "C:\Temp\Powershell-Logging"
    $LogFile = "$LogPath\MailboxPermissionLog.txt"
    if (!(Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath -Force | Out-Null }

    function Write-Log {
        param([string]$Message)
        $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Add-Content -Path $LogFile -Value "$Timestamp - $Message"
    }

    # Input Mailbox & Trustee
    $Identity = Read-Host "Enter the mailbox identity (user@domain.com)"
    $Trustee = Read-Host "Enter the trustee (user@domain.com)"
    $Tenant = (Get-OrganizationConfig).Name

    do {
        Clear-Host
        Write-Host "Mailbox Permissions Management Script" -ForegroundColor Cyan
        Write-Host "Tenant: $Tenant" -ForegroundColor Yellow
        Write-Host "Mailbox: $Identity" -ForegroundColor Yellow
        Write-Host "Trustee: $Trustee" -ForegroundColor Yellow
        Write-Host "======================================" -ForegroundColor Cyan
        Write-Host " 1. Add SendAs Permission"
        Write-Host " 2. Remove SendAs Permission"
        Write-Host " 3. Add FullAccess Permission"
        Write-Host " 4. Remove FullAccess Permission"
        Write-Host " 5. Add SendOnBehalf Permission"
        Write-Host " 6. Remove SendOnBehalf Permission"
        Write-Host " 7. List SendAs Permissions"
        Write-Host " 8. List FullAccess Permissions"
        Write-Host " 9. Add Calendar Permission"
        Write-Host "10. Remove Calendar Permission"
        Write-Host "11. List Calendar Permissions"
        Write-Host "12. List SendOnBehalf Permissions"
        Write-Host "13. List Mailbox Permissions"
        Write-Host "14. Add Recipient Permission"
        Write-Host "15. Remove Recipient Permission"
        Write-Host "16. List Recipient Permissions"
        Write-Host "17. List All Permissions"
        Write-Host "18. Update Primary UPN or Add Aliases"
        Write-Host "19. Exit"
        Write-Host "======================================" -ForegroundColor Cyan

        $choice = Read-Host "Choose an option (1-19)"

        switch ($choice) {
            1 {
                Add-RecipientPermission -Identity $Identity -Trustee $Trustee -AccessRights SendAs -Confirm:$false
                Write-Log "Added SendAs permission for $Trustee on $Identity"
            }
            2 {
                Remove-RecipientPermission -Identity $Identity -Trustee $Trustee -AccessRights SendAs -Confirm:$false
                Write-Log "Removed SendAs permission for $Trustee on $Identity"
            }
            3 {
                Add-MailboxPermission -Identity $Identity -User $Trustee -AccessRights FullAccess -AutoMapping:$false -Confirm:$false
                Write-Log "Added FullAccess permission for $Trustee on $Identity"
            }
            4 {
                Remove-MailboxPermission -Identity $Identity -User $Trustee -AccessRights FullAccess -Confirm:$false
                Write-Log "Removed FullAccess permission for $Trustee on $Identity"
            }
            5 {
                Set-Mailbox -Identity $Identity -GrantSendOnBehalfTo @{Add="$Trustee"}
                Write-Log "Added SendOnBehalf permission for $Trustee on $Identity"
            }
            6 {
                Set-Mailbox -Identity $Identity -GrantSendOnBehalfTo @{Remove="$Trustee"}
                Write-Log "Removed SendOnBehalf permission for $Trustee on $Identity"
            }
            7 {
                Get-RecipientPermission -Identity $Identity | Where-Object { $_.Trustee -eq $Trustee }
            }
            8 {
                Get-MailboxPermission -Identity $Identity | Where-Object { $_.User -eq $Trustee }
            }
            9 {
                $AccessRight = Read-Host "Enter calendar access right (AvailabilityOnly, LimitedDetails, Reviewer, Editor, etc.)"
                Add-MailboxFolderPermission -Identity "$Identity`:\Calendar" -User $Trustee -AccessRights $AccessRight -Confirm:$false
                Write-Log "Added Calendar ($AccessRight) permission for $Trustee on $Identity"
            }
            10 {
                Remove-MailboxFolderPermission -Identity "$Identity`:\Calendar" -User $Trustee -Confirm:$false
                Write-Log "Removed Calendar permission for $Trustee on $Identity"
            }
            11 {
                Get-MailboxFolderPermission -Identity "$Identity`:\Calendar"
            }
            12 {
                Get-Mailbox -Identity $Identity | Select-Object -ExpandProperty GrantSendOnBehalfTo
            }
            13 {
                Get-MailboxPermission -Identity $Identity
            }
            14 {
                Add-RecipientPermission -Identity $Identity -Trustee $Trustee -AccessRights SendAs -Confirm:$false
                Write-Log "Added Recipient permission (SendAs) for $Trustee on $Identity"
            }
            15 {
                Remove-RecipientPermission -Identity $Identity -Trustee $Trustee -AccessRights SendAs -Confirm:$false
                Write-Log "Removed Recipient permission (SendAs) for $Trustee on $Identity"
            }
            16 {
                Get-RecipientPermission -Identity $Identity
            }
            17 {
                Write-Host "=== SendAs Permissions ==="
                Get-RecipientPermission -Identity $Identity
                Write-Host "`n=== FullAccess Permissions ==="
                Get-MailboxPermission -Identity $Identity
                Write-Host "`n=== Calendar Permissions ==="
                Get-MailboxFolderPermission -Identity "$Identity`:\Calendar"
                Write-Host "`n=== SendOnBehalf Permissions ==="
                Get-Mailbox -Identity $Identity | Select-Object -ExpandProperty GrantSendOnBehalfTo
            }
            18 {
                Write-Host "1. Update Primary UPN"
                Write-Host "2. Add Aliases"
                $subChoice = Read-Host "Choose an option (1-2)"

                if ($subChoice -eq "1") {
                    Connect-MgGraph -Scopes "User.ReadWrite.All"
                    $NewUPN = Read-Host "Enter the new UPN (primary email)"
                    Update-MgUser -UserId $Identity -UserPrincipalName $NewUPN
                    Write-Log "Updated UPN for $Identity to $NewUPN"
                }
                elseif ($subChoice -eq "2") {
                    Connect-MgGraph -Scopes "User.ReadWrite.All"
                    $Aliases = Read-Host "Enter one or more aliases separated by commas"
                    $AliasesArray = $Aliases -split ","
                    $User = Get-MgUser -UserId $Identity
                    $CurrentAliases = $User.ProxyAddresses
                    foreach ($Alias in $AliasesArray) {
                        $Alias = $Alias.Trim()
                        if ($Alias -ne "") {
                            $CurrentAliases += "smtp:$Alias"
                            Write-Log "Added alias $Alias to $Identity"
                        }
                    }
                    Update-MgUser -UserId $Identity -ProxyAddresses $CurrentAliases
                }
            }
            19 {
                Write-Host "Exiting script..." -ForegroundColor Red
            }
            Default {
                Write-Host "Invalid choice, please select a valid option." -ForegroundColor Red
            }
        }
        Pause
    } until ($choice -eq 19)
}

# Run the menu
Set-MailboxPermissionsMenu
