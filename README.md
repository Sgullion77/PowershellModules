# Mailbox Permissions Management Script

#Author:   Seth Gullion
#Company:  Standard Office Systems of Atlanta

A **PowerShell script** to manage Exchange Online mailboxes, calendar permissions, and Microsoft Graph UPN/alias management from a single interactive menu. Designed for IT admins managing multiple tenants.

---

## Features

This script provides a menu-driven interface to:

- Add/remove mailbox permissions (`SendAs`, `FullAccess`, `SendOnBehalf`)
- Manage calendar permissions for individual users or bulk (CSV or comma-separated)
- Update mailbox identity, default trustee, and tenant connections
- Change primary UPN and add secondary aliases via Microsoft Graph
- List current UPN and aliases clearly marked as primary/secondary
- Log all actions to `C:\Temp\Powershell-Logging\MailboxPermissionLog.txt`

---

## Requirements

- **PowerShell 7.x or later**
- **Exchange Online Management Module** (`ExchangeOnlineManagement`)
- **Microsoft Graph PowerShell Module** (`Microsoft.Graph.Users`)
- Proper permissions in Exchange Online and Graph (`User.ReadWrite.All` for UPN/aliases)
- Network connectivity to tenant Exchange Online and Microsoft Graph

---

## Setup

1. Use IRM to run this script.
2. Open PowerShell as Administrator.
3. Install required modules (if not already installed):

```powershell
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Graph -Scope CurrentUser
Run the script:

open powershell as an admin and run:

irm https://raw.githubusercontent.com/Sgullion77/PowershellModules/refs/heads/main/Permissions.ps1 | iex

Enter your tenant name and default mailbox/trustee when prompted.

Menu Options
Option	Action
1	Add SendAs Permission – grants a user the ability to send email as the mailbox.
2	Add FullAccess Permission – grants a user full mailbox access without auto-mapping.
3	Add SendOnBehalf Permission – grants a user the ability to send on behalf of the mailbox.
4	Remove SendAs Permission – removes SendAs permission for the trustee.
5	Remove FullAccess Permission – removes FullAccess permission.
6	Remove SendOnBehalf Permission – removes SendOnBehalf permission.
7	Exit – closes the script.
8	List Current Permissions – lists SendAs, FullAccess, and SendOnBehalf permissions for the mailbox.
9	Add Calendar Permission (single trustee) – grant calendar access with a specific role (e.g., Reviewer, Editor).
10	Remove Calendar Permission (single trustee) – remove calendar access for a trustee.
11	List Calendar Permissions – shows current calendar access.
12	Bulk Add Calendar Permissions (comma-separated) – add calendar permissions for multiple users entered directly.
13	Bulk Add Calendar Permissions from CSV – add permissions for users listed in a CSV file (must include a User column).
14	Bulk Remove Calendar Permissions from CSV – remove calendar permissions for users in a CSV file.
15	Change Mailbox Identity – update the currently selected mailbox for management.
16	Change Default Trustee – change the default trustee email for permissions.
17	Change Tenant Name (and reconnect) – disconnects and reconnects to a new tenant; update mailbox and trustee.
18	Change Primary UPN and Add Aliases (Graph API) – update primary UPN and add one or more secondary aliases. Graph connection only occurs when this option is selected.
19	List Current UPN and Aliases (Graph API) – view mailbox UPN and secondary aliases in a table clearly labeled Primary/Secondary.

Logging
All changes are logged to:

C:\Temp\Powershell-Logging\MailboxPermissionLog.txt
Each log entry includes a timestamp, the executing user, and a description of the action.

Example Workflow
Run the script and connect to tenant dekalbhousing.org

Enter the mailbox mailbox@dekalbhousing.org(Mailbox that needs permissions added) and default trustee sosonsite@dekalbhousing(User that needs the mailbox permission.)

Add SendAs permission for Sosonsite → Option 1.

Grant calendar access to multiple users from CSV → Option 13.

Update primary UPN and add secondary aliases → Option 18.

Verify all aliases → Option 19.

Log all changes in C:\Temp\Powershell-Logging.

Usage Tips
Multiple trustees can be entered comma-separated or via CSV for bulk calendar permission updates.

Aliases can be updated multiple times; the script ensures duplicates are removed.

Graph API connection only triggers when needed, minimizing authentication prompts.

Contributions are welcome! You can:

Add more permission types

Integrate additional Graph API features

Improve logging or error handling

