# powershell-new-user
## New User Provisioning with PowerShell
Automates user provisioning in a hybrid Active Directory and Microsoft Entra ID environment, reducing manual onboarding time by approximately 90%

# What It Does
1. Reads in user alias (firstname.lastname) from a .txt file 
2. Prompts administrator to select the target Organizational Unit (OU) via an interactive menu
3. Generates a secure temporary password
4. Connects to on-premises Exchange and creates a remote mailbox, setting UPN, alias, display name, first name, last name, and password
5. Places the user in the correct AD OU and adds them to AD groups based on OU membership
6. Disconnects from on-premises Exchange
7. Sends an encrypted email to HR via Microsoft Graph containing the new user's credentials
8. Triggers a delta sync on the Entra Connect server to propagate changes to Entra ID
9. Connects to Exchange Online and applies mailbox policies (retention policy, photo change restriction)
10. Connects to Microsoft Graph and adds the user to cloud-based security groups
11. Disconnects from Microsoft Graph

# Why It Was Built
Manual provisioning of 1-3 users across on-premises Exchange, Active Directory, and Exchange Online required coordinating multiple admin consoles and waiting on sync cycles, taking 30–35 minutes. This script consolidates the entire workflow into a single execution, completing in roughly 3 minutes.

# Environment
- Hybrid AD and Entra ID (Azure AD Connect / Entra Connect)
- On-premises Exchange with remote mailbox provisioning
- Exchange Online
- Microsoft Graph API
- PowerShell

# Notes
Sensitive values (credentials, tenant IDs, etc.) are not included. The script is sanitized for portfolio use.
