<#
.SYNOPSIS
    This script automates the process of creating a new user account on the 2022 Exchange server 
.DESCRIPTION
    Reads from a txt file of names formatted as firstName.lastName, each on a new line, and sets new user properties
    based on the imported strings. Prompts user to select the OUs that the user will be created in, then generates a random
    password for the user before creating the account/mailbox. User account is automatically assigned to distribution and security
    groups based on the OU they were created in. Credentials for each account are printed to the terminal at the end of the script
    to allow for copying to an email for HR.
.EXAMPLE
    .\newHire.ps1
.NOTES
    Author: Isaac Aparicio
    Created: 08-11-2025
    Last Modified: 11-12-2025
    Version: 2.8
    Requires: PowerShell 5.1+ / 7+ (depending on your environment)
              Microsoft.Graph.*/Authentication version 2.25.0
              ExchangeOnlineManagement v3.5.1 (v3.9.x is incompatible) 
              Ability to remotely run an delta sync on the Entra ID connect server

    1.1: Added functionality that allows selection of new user's OU by integer input, rather than
         finishing the DN manually (ex: entering 1 then 3 instead of 'topOU/subOU')
    1.2: New user will be added to a different collection of groups based on top-level OU input
    1.3: Added secure password input and functionality to export new user credentials to a csv file
         * this functionality will be replaced later - replaced in v1.6 *
    1.4: Added output for newly created user's properties
    1.5: Added functionality to import new hire's name (John.Doe) from a csv file, and set different 
         account properties without needing user input
    1.6: Replaced prompt for secure password input with random, secure password generation that is
         printed with the associated UPN for all created users, once the csv file is fully processed
         * printing creds was for testing purposes and is not included in the final version *
    1.7: Read in the users with a .txt file instead of a .csv file
         * revisit this and convert back to .csv so columns/properties can be set and accessed by the script *
    1.8: Modified AD group assigning logic to use a hashtable instead of switches and if-statements 
    1.9: Updated password creation function (New-RandomPassword) to allow for clearer separation of 
         guaranteed and filler characters
    2.0: Broke out code into functions and added new ones (New-UserPassword, Select-OU, Add-UserToGroups, 
         New-UserMailbox, Connect-OnPremExchange, Disconnect-OnPremExchange)
    2.1: Added wait and check of user object before adding user to groups in AD
    2.2: Added Send-CredentialEmail function to send list of credentials in an email with PowerShell and Graph,
         which is auto-encrypted by a DLP sensitivity label/policy, replacing the process of copying the credentials
         that are printed out in the terminal
         * should implement a way to auto-apply encryption with the DLP policy dependency *
    2.3: Added Import-PSsession logic which imports the cmdlets used by Exchange Management Shell on the on-prem Exchange 
         server
    2.4: Adjusted the wait duration when checking the AD user object exists before adding to AD groups, to
         5 seconds from 10 seconds
    2.5: Implemented Azure AD Sync function that is ran after creating all users
    2.6: Implemented Set-MailboxPolicy function
    2.7: Addded Add-EntraGroupMember function
            - Adds user to cloud based groups via Graph API
    2.8: Added comments, and comment blocks for each function, and modified the order of the functions to match the 
         flow of the provisioning process.
    
    # --- Function Index ---
    # 1. Connect-OnPremExchange    -> Establishes PSSession with on-prem Exchange server/EMS
    # 2. Select-OU                 -> Prompts for OU selection interactively
    # 3. New-UserPassword          -> Generates secure passwords
    # 4. New-UserMailbox           -> Creates a remote mailbox in on-prem Exchange
    # 5. Add-UserToGroups          -> Assigns AD groups based on OU
    # 6. Disconnect-OnPremExchange -> Closes on-prem Exchange session
    # 7. Send-CredentialEmail      -> Sends new user credentials securely via Graph API
    # 8. Start-AzureADSync         -> Invokes delta sync on AAD Connector server
    # 9. Set-MailboxPolicy         -> Applies custom mailbox/retention policies
    # 10. Add-EntraGroupMember     -> Adds user to Entra group via Graph API

    ! --- Known Issues --- !
    - Start-AzureADSync does not properly print out all the specified properties when the sync has completed,
      only the "success" status message
    - MissingMethodException: WithBroker(...) method not found
      Root cause: conflict with Microsoft.Identity.Client dependency due to incompatible version used by ExchangeOnlineManagement
      module v3.9x.
      Additionally, multiple versions of PackageManagement/PowerShellGet caused module loading inconsistencies.
      Resoultion:
       - Forced correct module versions in session:
          ~ PowerShellGet v2.2.5
          ~ PackageManagement v1.4.8.1
       - Installed stable ExchangeOnlineManagement version:
          ~ v3.5.1 (known working version as of 4/9/26)
       - Installed using:
           Install-Module ExchangeOnlineManagement -RequiredVersion 3.5.1 -Scope Current User -Force -AllowClobber

    ! --- Future Functionalities --- !
    - Possibly include user's destination parent and child OU (FTE, Sales), in the txt/csv file to allow for greater automation,
      instead of having to input the integers that correspond to the OUs
        + ex: userList.txt line 1: "John.Doe,FTE,Sales"
        + ex: userList.csv row 1: Name | ParentOU | SubOU
                           row 2: John.Doe | FTE | Sales      
    - Remove all test-based Write-Output lines to prepare it for full prod use
    - Clean up and add comments
#>

# --- Connect to on-prem Exchange remotely ---
<#
.SYNOPSIS
    Establishes a PSSession with the on-prem Exchange server/EMS
.DESCRIPTION
    Creates a new PSSession with the on-Prem Exchange server and imports EMS commands so that 
    New-RemoteMailbox (EMS native command), can be ran outside of the EMS
.PARAMETER ServerFQDN
    String of the on-prem Exchange server
.PARAMETER ConfigurationName
    String of the targeted configuration name when connecting to the on-prem Exchange server
.OUTPUTS
    N/A
.EXAMPLE
    N/A
.NOTES
    Called in the main before processing the list of user objects to create
#>
function Connect-OnPremExchange {
    param(
        [string]$ServerFQDN = "YourExchangeServer.localdomain",
        [string]$ConfigurationName = "Microsoft.Exchange"
    )

    Write-Host "Connecting to on-prem Exchange..." -ForegroundColor Cyan

    try {
        $global:OnPremSession = New-PSSession -ConfigurationName $ConfigurationName `
            -ConnectionUri http://$ServerFQDN/PowerShell/ `
            -Authentication Kerberos -ErrorAction Stop

        Import-PSSession $OnPremSession -DisableNameChecking -ErrorAction Stop | Out-Null
        Write-Host "Connected to on-prem Exchange" -ForegroundColor Green
    }
    catch {
        if ($global:OnPremSession) {
            Remove-PSSession $OnPremSession -ErrorAction SilentlyContinue
            Remove-Variable OnPremSession -Scope Global -ErrorAction SilentlyContinue
        }
        Write-Error "Failed to connect to on-prem Exchange. $_"
    }
}

# --- Select target OU for the new user ---
<#
.SYNOPSIS
    Prompts user for OU selection for the new user
.DESCRIPTION
    Prints numbered options for top/sub-level OUs where the new user object will be located, and accepts 
    valid user input of an integer that corresponds to the desired OU
.PARAMETER baseOU
    String for building the users destination OU
.OUTPUTS
    PSCustomObject with the top/sub-level OU names, and full destination OU name properties
.EXAMPLE
    $finalOU = Select-OU -baseOU $baseOU
    $finalOU.DistinguishedName
.NOTES
    Called in the main for each new user to select their destination OU
#>
function Select-OU {
    param(
        [string]$baseOU # ex: "OU=Users,OU=Corporate,DC=localdomain"
    )

    # Top-level OUs nested in "OU=Users"
    $topLevelOUs = @(
        @{Name='SampleOU1'},
        @{Name='SampleOU2'},
        @{Name='SampleOU3'}
    )

    # prints top-level OU options
    Write-Host "Select the top-level OU:"
    for($i=0; $i -lt $topLevelOUs.Count; $i++) {
        Write-Host "$($i+1). $($topLevelOUs[$i].Name)"
    }

    # prompts user to enter integer that corresponds to desired top-level OU, and checks for valid input via a loop
    do {
        $topChoice = Read-Host "Enter choice (1-$($topLevelOUs.count))"
        if ($topChoice -match '^\d+$') {
            $topChoiceInt = [int]$topChoice
            $isValid = $topChoiceInt -ge 1 -and $topChoiceInt -le $topLevelOUs.Count
        } else {
            $isValid = $false
        }
        if(-not $isValid) {
            Write-Host "Invalid selection. Please choose a number from 1 to $($topLevelOUs.Count)" -ForegroundColor Red
        }
    } until ($isValid)

    $selectedTopOU = $topLevelOUs[$topChoice -1] # Subtract by 1 to match the human-friendly interger entered to the array's index
    $selectedTopOUDN = "OU=$($selectedTopOU.Name),$baseOU" # Create full OU path for retrieving it's sub OUs. ex: "OU=SampleOU1" + "OU=Users,OU=Corporate,DC=localdomain"

    # Get sub-OUs dynamically from AD
    $subOUs = Get-ADOrganizationalUnit -SearchBase $selectedTopOUDN -Filter * -SearchScope OneLevel | Sort-Object Name

    # prints sub-OU options for the selected top-level OU
    Write-Host "`nSelect sub-OU under " -NoNewLine
    Write-Host "$($selectedTopOU.name)" -ForegroundColor Cyan
    $subOUs += [PSCustomObject]@{
        Name = "SampleOU1 (no sub OU)"
        DistinguishedName = $null
    }
    for ($j=0; $j -lt $subOUs.Count; $j++) {
        Write-Host "$($j+1). $($subOUs[$j].name)"
    }

    # prompts user to enter integer that corresponds to desired top-level OU, and checks for valid input via a loop
    do {
        $subChoice = Read-Host "Enter choice (1-$($subOUs.Count))"
        if ($subChoice -match '^\d+$') {
            $subChoiceInt = [int]$subChoice
            $isValid = $subChoiceInt -ge 1 -and $subChoiceInt -le $subOUs.Count
        }
        if (-not $isValid) {
            Write-Host "Invalid selection. Please choose a number from 1 to $($subOUs.Count)." -ForegroundColor Red
        }
    } until ($isValid)

    # Subtract by 1 to match the human-friendly number to the array's index
    $selectedSubOU = $subOUs[$subChoice - 1]
    
    # Build the final OU DN string
    if ($null -eq $selectedSubOU.DistinguishedName) {
        $selectedSubOU.Name = "N/A"
        $finalOU = $selectedTopOUDN
    } else {
        $finalOU = $selectedSubOU.DistinguishedName
    }
    return [PSCustomObject]@{
        TopOU = $selectedTopOU.Name
        SubOU = $selectedSubOU.Name
        DistinguishedName = $finalOU
    }
}

# --- Generates and stores a password ---
<#
.SYNOPSIS
    Generates a random, secure password and returns both plain and secure versions.
.DESCRIPTION
    Uses a combination of uppercase, lowercase, numeric, and special characters 
    to generate a password that meets complexity requirements. Ensures at least 1
    upper char, lower char, and digit, exactly 2 special chars, and randomizes the rest (excluding special).
.PARAMETER length
    The total length of the password to generate (default: 10).
.OUTPUTS
    PSCustomObject with PlainPassword and SecurePassword properties.
.EXAMPLE
    $pw = New-UserPassword -length 12
    $pw.PlainPassword
.NOTES
    Called in the main loop for each new user to generate their unique credentials.
#>
function New-UserPassword {
    param([int]$length = 10)

    # Character groups
    $upper   = ([char[]]([char]'A'..[char]'Z') | Where-Object {$_ -notin @('I','O')})
    $lower   = ([char[]]([char]'a'..[char]'z') | Where-Object {$_ -notin @('i','l','o')})
    $digits  = ([char[]]([char]'2'..[char]'9'))
    $special = @('!','@','$','-','+','?','^','&','*')

    # Allowed non-special characters
    $allowedChars = $upper + $lower + $digits 

    # Ensure at least 1 upper, lower, and digit
    $requiredChars = @(
        ($upper | Get-Random),
        ($lower | Get-Random),
        ($digits | Get-Random)
    )

    # Add up to 2 special characters
    $specialChars = $special | Get-Random -Count 2
    $requiredChars += $specialChars

    # Fill remaining length with non-special characters
    $remainingChars = $length - $requiredChars.Count
    $randomChars = 1..$remainingChars | ForEach-Object { $allowedChars | Get-Random }

    # Combine and shuffle
    $plainPassword = -join ($requiredChars + $randomChars | Get-Random -Count $length)

    # Convert to secure string for mailbox creation
    $securePassword = ConvertTo-SecureString $plainPassword -AsPlainText -Force 

    # Return both passwords as a PSCustomObject
    return [PSCustomObject]@{
        PlainPassword = $plainPassword
        SecurePassword = $securePassword
    }
}

# --- Create the new mailbox ---
<#
.SYNOPSIS
    Creates a new remote mailbox/user object in AD
.DESCRIPTION
    Runs the EMS-based New-RemoteMailbox command with previously set properties
.PARAMETER displayName
    String for the user's full name
.PARAMETER firstName
    String for the user's first name
.PARAMETER lastName
    String for the user's last name
.PARAMETER alias
    String for the user alias
.PARAMETER userPrincipalName
    String for the UPN
.PARAMETER remoteRoutingAddress
    String for the onmicrosoft.com address
.PARAMETER onPremisesOrganizationalUnit
    String of the destination OU created with Select-OU
.PARAMETER securePassword
    Secure string of password created with New-UserPassword
.OUTPUTS
    N/A
.EXAMPLE
    
.NOTES
    Called in the main loop for each new user to create their mailbox/account
#>
function New-UserMailbox {
    param(
        [string]$displayName, 
        [string]$firstName, 
        [string]$lastName, 
        [string]$alias, 
        [string]$userPrincipalName,
        [string]$remoteRoutingAddress, 
        [string]$onPremisesOrganizationalUnit,
        [securestring]$securePassword
    )
    # --- Create the remote mailbox/user ---
    try {
        New-RemoteMailbox `
            -Name $displayName `
            -DisplayName $displayName `
            -FirstName $firstName `
            -LastName $lastName `
            -Alias $alias `
            -UserPrincipalName $userPrincipalName `
            -RemoteRoutingAddress $remoteRoutingAddress `
            -OnPremisesOrganizationalUnit $onPremisesOrganizationalUnit `
            -Password $securePassword `
            -ErrorAction Stop | Out-Null

        Write-Host "Mailbox created for $displayName" -ForegroundColor Green
    } 
    catch {
        Write-Warning "Mailbox creation failed $_"
        return
    } 
}

# --- Add user to the appropriate AD groups ---
<#
.SYNOPSIS
    Adds the user(s) to groups in AD based on their created destination OU
.DESCRIPTION
    Uses a nested hashtable to determine which groups the user(s) needs to be added to, based on
    previously selected top and sub-level OUs. Removes duplicate OUs, then adds the user to the available
    group(s)
.PARAMETER topOU
    String of the top-OU selected with Select-OU
.PARAMETER subOU
    String of the sub-OU selected with Select-OU
.PARAMETER adUser
    Microsoft.ActiveDirectory.Management.ADUser object that contains the entire AD user object so properties
    like DistinguishedName can be accessed when adding the user to groups
.OUTPUTS
    N/A
.EXAMPLE
    
.NOTES
    Called in the main loop for each user, within a while-loop that checks that the account exists before attempting 
    to add them to groups
#>
function Add-UserToGroups {
    param(
        [string]$topOU,
        [string]$subOU,
        [Microsoft.ActiveDirectory.Management.ADUser]$adUser # <-- AD object
    )
    
     # Full time employee groups (common groups)
    $fteGroups = @(
        "FTE Group1"
        "FTE Group2"
        "FTE Group3"
    )

    # Define your OU-based adgroup table
    $ouGroupTable = @{
        'SampleOU1' = @{
            'SubOU1' = @("SubOU1-All")
            'SubOU2' = @("SubOU2-ALL", "SampleSecurityGroup") # Special PSG groups
            'SubOU3' = @("SubOU3-All")
        }
        'SampleOU2' = @{}
        'SampleOU3'   = @{
            '*'         = @() # Default SampleOU3. Used when a sub OU is selected that isn't declared in the hashtable
            'SubOU1'   = @(
                "SampleGroup1", 
                "SampleGroup2", 
                "SampleGroup3", 
                "SampleGroup4"
            )
            'SubOU2' = @("SampleGroup", "SampleGroup2")
            'SubOU3' = @("SampleGroup","SampleGroup") 
        }
    }

    $adGroups = @()

    # Add sample groups for certain top-level OUs
    if ($topOU -in @('SampleOU1', 'SampleOU2')) {
        $adGroups += $fteGroups
    }

    # Add top-level OU-specific groups
    if ($ouGroupTable.ContainsKey($topOU)) {
        $topOUEntry = $ouGroupTable[$topOU]
        
        if($topOUEntry -is [hashtable]) {
            if($topOUEntry.ContainsKey($subOU)) {
                $adGroups += $topOUEntry[$subOU]
            } elseif ($topOUEntry.ContainsKey('*')) {
                $adGroups += $topOUEntry['*']
            }
        } elseif ($topOUEntry -is [array]) {
            $adGroups += $topOUEntry
        }
    }

    # --- Remove duplicates ---
    $adGroups = $adGroups | Sort-Object -Unique

    # --- pull and add groups ---
    Write-Host "User will be added to groups:" -ForegroundColor Cyan
    $adGroups | ForEach-Object { Write-Host "- $_" }

     ForEach ($groupName in $adGroups) {
        $adgroup = Get-ADGroup -Filter "Name -eq '$groupName'"

        if ($adgroup) {
            try {
                Add-ADGroupMember -Identity $adgroup.DistinguishedName -Members $adUser.DistinguishedName -ErrorAction Stop
                Write-Host "Added $($adUser.SamAccountName) to $($adgroup.Name)"
            } catch {
                Write-Warning "Failed to add $($adUser.SamAccountName)  to $($adgroup.Name). Error $_"
            }
        }
        else {
            Write-Warning "Group not found: $groupName"
        }   
    }
}

# --- Clean up on-prem session ---
<#
.SYNOPSIS
    Disconnects the user from the on-prem Exchange session that was established with Connect-OnPremExchange
.DESCRIPTION
    Disconnects the user from the on-prem Exchange session to free up system resources, avoid session
    exhaustion and conflicts, and closes authentication channels
.PARAMETER <parameter name>
    N/A
.OUTPUTS
    N/A
.EXAMPLE
    
.NOTES
    Called in the main after each user has been created and added to eligible AD groups
#>
function Disconnect-OnPremExchange {
    if($global:OnPremSession) {
        Write-Host "Closing on-prem Exchange session..." -ForegroundColor Yellow
        Remove-PSSession $OnPremSession
        Remove-Variable OnPremSession -Scope Global -ErrorAction SilentlyContinue
        Write-Host "On-prem Exchange session closed" -ForegroundColor Green
    }
    else {
        Write-Host "No active on-prem Exchange session found" -ForegroundColor DarkGray
    }
}

# --- Send user credentials in an encrypted email using Graph ---
<#
.SYNOPSIS
    Sends an email with the user credentials via Graph API
.DESCRIPTION
    Sends an email with a table of the newly created account UPNs + their password
.PARAMETER CredentialList
    PSCustomObject of user UPNs and passwords
.PARAMETER Recipients
    Array of strings that contain the recipients of the credential email
.OUTPUTS
    N/A
.EXAMPLE
    
.NOTES
    Called in the main after all users have been created and the on-prem Exchange session has disconnceted
#>
function Send-CredentialEmail {
    param (
        [pscustomobject]$CredentialList,
        [string[]]$Recipients # Accepts 1 or more recipients
    )

    # --- Connect to Microsoft Graph --- #
    if(-not (Get-MgContext)) {
        Connect-MgGraph -Scopes "Mail.Send","Mail.ReadWrite","User.Read" | Out-Null
    }

    $senderEmail = (get-mgcontext).Account

    # Build recipient object for graph
    $recipientEmails = @()
    foreach ($r in $Recipients) {
        $recipientEmails += @{
             EmailAddress = @{ 
                Address = $r 
            } 
        }
    }

    # Format credentials as plain text
    $credentialText = $credentialList | ForEach-Object {
        "Username: $($_.UserPrincipalName)`nPassword:  $($_.Password)`n"
    } | Out-String

    try {
        Send-MgUserMail -UserId $senderEmail -Message @{
            Subject = "New Hire Credentials"
            Body    = @{
                ContentType = "Text"
                Content = "Hi,`n`nPlease see below:`n`n$credentialText" + "This email was sent via PowerShell and Microsoft Graph API"
            }
            ToRecipients = $recipientEmails
        } -SaveToSentItems -ErrorAction Stop

        Write-Host "Credentials email sent successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "An error occurred while sending the email:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Yellow
    }
}

# --- Runs an Azure AD Sync against the AAD connector server
<#
.SYNOPSIS
    Runs a delta sync on the AAD Connector server
.DESCRIPTION
    Delta sync is ran and the script wait til the sync is complete before continuing with the remaining
    functions which require the user/mailbox to be synced to Entra. The sync status check will timeout after
    5 minutes, exiting the script, so the remaining functions will not execute
.PARAMETER ComputerName
    String with the server name where the AAD Connector is located
.PARAMETER CheckInterval
    Int representing the interval of time (in seconds) of how often the sync status will be checked 
    (default=5)
.PARAMETER TimeoutMinutes
    Int representing the max amount of time (in minutes) spent checking on the sync status until the sync is considered
    to have "timed out" (default=5)
.OUTPUTS
    N/A
.EXAMPLE
    
.NOTES
    Called in the main before modifying cloud-based user properties in Entra, and mailbox policies in
    Exchange Online
#>
function Start-AzureADSync {
    param (
        [string]$ComputerName = "YourEntraConnectServer.localdomain",
        [int]$CheckInterval = 5,
        [int]$TimeoutMinutes = 5
    )
    Write-Host "Starting Azure AD Sync..."

    $syncResult = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        param($CheckInterval,$TimeoutMinutes)

        Start-ADSyncSyncCycle -PolicyType Delta
        $startTime = Get-Date

        do {
            $syncStatus = (Get-ADSyncScheduler).SyncCycleInProgress
            if($syncStatus -eq $true) {
                Write-Progress "Sync still running...checking again in $CheckInterval seconds."
                Start-Sleep -Seconds $CheckInterval
            }

            if((Get-Date) -gt $startTime.AddMinutes($TimeoutMinutes)) {
                throw "Azure AD Sync timed out after $TimeoutMinutes minutes."
            }
        } while ($syncStatus -eq $true)

        Write-Host "AD Sync complete. Result:"
    } -ArgumentList $CheckInterval, $TimeoutMinutes

    $syncResult | Select-Object Result, Type, StartTime, EndTime | Format-Table -AutoSize | Out-String | Write-Host -ForegroundColor Green
}

# --- Sets Role Assignment Policy: Users Cannot Change Their Password Policy, and Retention Policy: 5 days Deleted Items
<#
.SYNOPSIS
    Sets custom policies on the newly created mailboxes in Exchange Online
.DESCRIPTION
    While connected to Exchange Online, waits till the user/mailbox to be found, then sets the following
    mailbox policies:
    Role Assignment Policy: CustomRolePolicy
    Retention Policy:       CustomRetentionPolicy
.PARAMETER userPrincipalName
    Array of strings containing the UPNs for the newly created user accounts
.OUTPUTS
    N/A
.EXAMPLE
    
.NOTES
    Called in the main after a sync is ran and the session is connected to Exchange Online
#>
function Set-MailboxPolicy {
    param (
        [string[]]$userPrincipalName
    )

    Write-Host "Connecting to Exchange Online - " -ForegroundColor Blue

    foreach ($upn in $userPrincipalName) {
        $maxAttempts = 12 # wait up to ~2 minutes
        $attempt = 1
        $userFound = $false

        Write-Host "Waiting for user to appear in Exchange Online..."
        do {
            try {
                $mailboxCheck = Get-EXOMailbox -Identity $upn -ErrorAction Stop
                if($mailboxCheck) {
                    Write-Host "$upn found in Exchange Online after $attempt attempts"
                    $userFound = $true
                }
            }
            catch {
                Write-Host "$upn not found yet in Exchange Online... (Attempt $attempt)" -ForegroundColor Yellow
                Start-Sleep -Seconds 10
            }
            $attempt++
        } until ($userFound -or $attempt -gt $maxAttempts)

        if (-not $userFound) {
            Write-Warning "User not found in Exchange online after waiting period. Skipping mailbox policy."
        }

        try {
            Set-Mailbox -Identity $upn `
                -RoleAssignmentPolicy "CustomRolePolicy" `
                -RetentionPolicy "CustomRetentionPolicy"
        
            $mailbox = Get-Mailbox -Identity $upn | 
                Select-Object UserPrincipalName, RoleAssignmentPolicy, RetentionPolicy
        
            Write-Host "Applied policies for $upn"
            $mailbox | Format-Table -AutoSize | Out-String | Write-Host
        }
        catch {
                Write-Warning "Failed to apply mailbox policies for $upn. $_"
        }
    }
}

# --- Adds user(s) cloud-based groups via Graph API
<#
.SYNOPSIS
    Adds newly created user(s) to Entra-based groups via Graph API
.DESCRIPTION
    Using the credentialList variable's UserPrincipalName property, adds each available user to cloud-based groups. 
    The user ID needed for Graph API operations is obtained using the UPN,
    while the group ID is obtained using the fixed group name string when using Get-MgGroup
.PARAMETER UserPrincipalNames
    Array of strings that contain the UPN(s) for the newly created user(s)
.OUTPUTS
    N/A
.EXAMPLE
    
.NOTES
    Called in the main after a sync is ran and has completed successfully
#>
function Add-EntraGroupMember {
    param (
        [string[]]$UserPrincipalNames
    )

    Write-Host "Connecting to Microsoft Graph" -ForegroundColor Blue

    Connect-MgGraph -Scopes "Group.ReadWrite.All", "User.Read.All" | Out-Null
    
    # --- Resolve group by display name ---
    $group = Get-MgGroup -Filter "displayName eq 'CloudGroup'"
    $groupId = $group.Id
    
    # --- Get all UPNs from newUserCredentials
    # $allUPNs = $newUserCredentials | Select-Object -ExpandProperty UserPrincipalName
    
    foreach ($upn in $UserPrincipalNames) {
        try {
            $userId = (Get-MgUser -UserId $upn).Id
            New-MgGroupMember -GroupId $groupId -DirectoryObjectId $userId -ErrorAction Stop
            Write-Host "Added $upn to "$($group.DisplayName)" security group"
        }
        catch {
            Write-Warning "Failed to add $upn to group. $_"
        }
    }

    Write-Host "Disconnecting from Microsoft Graph" -ForegroundColor Blue

    Disconnect-MgGraph | Out-Null
}

# --- Main ---
$domainName = "localdomain.com"
$baseOU = "OU=Users,OU=Corporate,DC=localdomain" # Base DN, ex: your root domain
$remoteRoutingDomain = "localdomain.mail.onmicrosoft.com"
$txtFilePath = "C:\New User Script\userList.txt"
$credentialList = @()
$recipients = @("HR.Email@localdomain.com")

if (-not (Test-Path $txtFilePath)) {
    Write-Host "Cannot find path '$txtFilePath' because it does not exist" -ForegroundColor Red
    exit
}

$users = Get-Content -Path $txtFilePath

Connect-OnPremExchange

foreach ($user in $users) {
    $firstName, $lastName        = $user -split '\.'
    $displayName                 = "$firstName $lastName"                       # John Doe
    $userPrincipalName           = "$firstName.$lastName@$domainName"           # John.Doe@localdomain.com
    $remoteRoutingAddress        = "$firstName.$lastName@$remoteRoutingDomain"  # John.Doe@localdomain.mail.onmicrosoft.com
    
    Write-Host "----- Starting $displayName -----" -ForegroundColor Magenta 

    $finalOU = Select-OU -baseOU $baseOU
    $passwords = New-UserPassword -length 10
    New-UserMailbox -DisplayName $displayName `
                      -FirstName $firstName `
                      -LastName $lastName `
                      -Alias "$firstName.$lastName" `
                      -UserPrincipalName $userPrincipalName `
                      -RemoteRoutingAddress $remoteRoutingAddress `
                      -OnPremisesOrganizationalUnit $finalOU.DistinguishedName `
                      -SecurePassword $passwords.SecurePassword

    # Try up to 5 times with a 5s delay
    $maxAttempts = 5
    $attempt = 1
    $adUser = $null

    while (-not $adUser -and $attempt -le $maxAttempts) {
        $adUser = Get-ADUser -Filter { UserPrincipalName -eq $userPrincipalName } -ErrorAction SilentlyContinue

        if ($adUser) {
            Write-Host "Found $userPrincipalName in AD on attempt $attempt"
            Add-UserToGroups -topOU $finalOU.TopOU -subOU $finalOU.SubOU -adUser $adUser
            break
        }
        else {
            Write-Host "$userPrincipalName not found in AD yet. Waiting 5 seconds... (Attempt $attempt of $maxAttempts)" -ForegroundColor Yellow
            Start-Sleep -Seconds 5
            $attempt++
        }
    }

    if (-not $adUser) {
        Write-Warning "User $userPrincipalName could not be found in AD after $maxAttempts attempts. Group assignment skipped."
    }

    # Collect plain password for later
    $credentialList += [PSCustomObject]@{
        UserPrincipalName = $userPrincipalName
        Password = $passwords.PlainPassword
    }

    Write-Host "----- Finished $displayName -----" -ForegroundColor Magenta 
}

Disconnect-OnPremExchange

Send-CredentialEmail -CredentialList $credentialList -Recipients $recipients

Start-AzureADSync

Connect-ExchangeOnline -ShowBanner:$false # weird cmdlet where the login window pops-up in background - look into making it pop into foreground
Set-MailboxPolicy -userPrincipalName $credentialList.UserPrincipalName
Disconnect-ExchangeOnline -Confirm:$false

Add-EntraGroupMember -UserPrincipalNames $credentialList.UserPrincipalName

Write-Host "`n--- Summary --- "
$credentialList.UserPrincipalName | Format-Table -AutoSize | Out-String | Write-Host
Write-Host "Total users processed: $($credentialList.Count)" -ForegroundColor Green
