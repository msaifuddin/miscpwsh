# Required Modules
# Automatically install specific Microsoft.Graph modules if not installed
$requiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.DeviceManagement",
    "Microsoft.Graph.Identity.DirectoryManagement"
)

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "$module module not found. Installing..." -ForegroundColor Yellow
        try {
            Install-Module -Name $module -Scope CurrentUser -Force -ErrorAction Stop
            Write-Host "$module module installed successfully." -ForegroundColor Green
        } catch {
            Write-Host "Failed to install $module module. Please install it manually." -ForegroundColor Red
            exit
        }
    } else {
        Write-Host "$module module is already installed." -ForegroundColor Green
    }
}

# Import necessary modules
foreach ($module in $requiredModules) {
    try {
        Import-Module $module -ErrorAction Stop
        Write-Host "Imported $module module." -ForegroundColor Green
    } catch {
        Write-Host "Failed to import $module module. Please check the installation." -ForegroundColor Red
        exit
    }
}

# Parameters
$clientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e" # I'm identified as Microsoft Graph PowerShell SDK
$graphScopes = @(
    "User.Read.All",
    "DeviceManagementManagedDevices.Read.All",
    "RoleManagement.ReadWrite.Directory"
)

# Global variables
$global:userResults = @()
$global:deviceResults = @()
$global:isConnected = $false
$global:userDisplayName = ""

# Pagination settings
$global:pageSize = 10
$global:currentPage = 1
$global:totalPages = 1

# Helper function to read a single key including Enter
function Read-SingleKey {
    $keyInfo = [System.Console]::ReadKey($true)
    if ($keyInfo.Key -eq 'Enter') {
        return 'ENTER'
    } else {
        return $keyInfo.KeyChar.ToString().ToUpper()
    }
}

# Function to display simplified error messages
function Show-SimpleErrorMessage {
    Write-Host "NOT FOUND" -ForegroundColor Red
}

# Function to display simplified not found messages
function Show-NotFoundMessage {
    param (
        [string]$Context
    )
    if ($Context -eq 'Users') {
        Write-Host "NO USERS FOUND" -ForegroundColor Yellow
    } elseif ($Context -eq 'Devices') {
        Write-Host "NO DEVICES FOUND" -ForegroundColor Yellow
    } else {
        Write-Host "NOT FOUND" -ForegroundColor Yellow
    }
}

# Function to Draw a Consistent Banner
function Draw-Banner {
    param (
        [string]$ConnectedUser = "",
        [string]$Version = "v1.2"
    )

    Clear-Host

    # Old-school ASCII Art Banner
    $bannerWidth = 50
    $bannerText = "IT Support Helper"
    $padding = ($bannerWidth - 2 - $bannerText.Length) / 2
    $paddedText = ' ' * [math]::Floor($padding) + $bannerText + ' ' * [math]::Ceiling($padding)

    $topBottomBorder = '+' + '-' * ($bannerWidth - 2) + '+'
    $emptyLine = '|' + ' ' * ($bannerWidth - 2) + '|'

    Write-Host $topBottomBorder -ForegroundColor Cyan
    Write-Host $emptyLine -ForegroundColor Cyan
    Write-Host "|$paddedText|" -ForegroundColor Cyan
    Write-Host $emptyLine -ForegroundColor Cyan

    # Ensure version fits at the bottom-right
    $spaceForVersion = $bannerWidth - 3 - $Version.Length
    $versionLine = '|' + ' ' * $spaceForVersion + $Version + ' |'

    Write-Host $versionLine -ForegroundColor Cyan
    Write-Host $topBottomBorder -ForegroundColor Cyan

    if (-not [string]::IsNullOrEmpty($ConnectedUser)) {
        Write-Host ""
        Write-Host "Connected as: $ConnectedUser" -ForegroundColor Green
    }

    Write-Host ""
}

# Authenticate and initialize graph client
function Authenticate-Graph {
    Write-Host "`nAuthenticating to Azure AD..." -ForegroundColor Cyan

    try {
        # Connect to Microsoft Graph with interactive login and suppress welcome message
        Connect-MgGraph -ClientId $clientId -Scopes $graphScopes -NoWelcome -ErrorAction Stop

        # Get the current authentication context
        $context = Get-MgContext

        if ($context -and $context.Account) {
            $global:userDisplayName = $context.Account
            $global:isConnected = $true
            Write-Host "Authenticated as $($global:userDisplayName)" -ForegroundColor Green
            # Pause to allow user to see the message before proceeding
            Read-Host "Press Enter to continue..."
            # Draw the banner with connected user
            Draw-Banner $global:userDisplayName
        } else {
            Write-Host "Authentication failed or user information is unavailable." -ForegroundColor Red
            $global:isConnected = $false
            # Pause to allow user to see the failure message
            Write-Host "Press Enter to retry or 'E' to exit." -ForegroundColor Yellow
            $retryChoice = Read-SingleKey
            if ($retryChoice -eq 'E') {
                Write-Host "Goodbye!" -ForegroundColor Green
                exit
            }
        }

    } catch {
        Show-SimpleErrorMessage
        $global:isConnected = $false
        # Pause to allow user to see the error message
        Write-Host "Press Enter to retry or 'E' to exit." -ForegroundColor Yellow
        $retryChoice = Read-SingleKey
        if ($retryChoice -eq 'E') {
            Write-Host "Goodbye!" -ForegroundColor Green
            exit
        }
    }
}

# Function to activate/deactivate PIM roles
function Activate-PIM {
    param()
    
    # Function to read user input and detect special keys
    function Read-UserInput {
        param(
            [string[]]$SpecialKeys = @()
        )
        
        $input = ''
        while ($true) {
            $key = [System.Console]::ReadKey($true)
            if ($key.Key -eq 'Escape') {
                throw 'UserCancelled'
            } elseif ($SpecialKeys -contains $key.KeyChar.ToString().ToUpper()) {
                throw $key.KeyChar.ToString().ToUpper()
            } elseif ($key.Key -eq 'Enter') {
                Write-Host  # Move to the next line
                break
            } elseif ($key.Key -eq 'Backspace') {
                if ($input.Length -gt 0) {
                    $input = $input.Substring(0, $input.Length - 1)
                    [System.Console]::Write("`b `b")  # Move cursor back, overwrite character, move back again
                }
            } else {
                $input += $key.KeyChar
                [System.Console]::Write($key.KeyChar)
            }
        }
        return $input
    }

    if (-not $global:isConnected) {
        Write-Host "Please connect first." -ForegroundColor Yellow
        Start-Sleep -Seconds 1
        return
    }

    while ($true) {
        try {
            Draw-Banner $global:userDisplayName
            Write-Host "Loading PIM roles..." -ForegroundColor Cyan

            # Get current user information
            $currentUser = Get-MgUser -UserId $global:userDisplayName
            $currentUserId = $currentUser.Id

            # Retrieve eligible roles for the user
            $eligibleRoles = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -Filter "principalId eq '$currentUserId'" -ExpandProperty RoleDefinition

            if (-not $eligibleRoles) {
                Write-Host "No eligible roles found for your account." -ForegroundColor Yellow
                Write-Host "`nPress Enter to return to the main menu." -ForegroundColor Yellow
                [void][System.Console]::ReadKey($true)
                return
            }

            # Retrieve active role assignment schedule instances for the user
            $activeRoles = Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance -Filter "principalId eq '$currentUserId'" -ExpandProperty RoleDefinition
            $activeRoleIds = $activeRoles | Select-Object -ExpandProperty RoleDefinition | Select-Object -ExpandProperty Id

            # Display available roles with index for easy selection
            Clear-Host
            Draw-Banner $global:userDisplayName
            Write-Host "`nEligible Roles:" -ForegroundColor Cyan
            $index = 1
            $roleIndexMap = @{}
            foreach ($role in $eligibleRoles) {
                $roleId = $role.RoleDefinition.Id
                $roleDisplayName = $role.RoleDefinition.DisplayName

                if ($activeRoleIds -contains $roleId) {
                    Write-Host "$index. [Active] $roleDisplayName" -ForegroundColor Green
                } else {
                    Write-Host "$index. $roleDisplayName" -ForegroundColor White
                }
                $roleIndexMap[$index] = $role
                $index++
            }

            # Prompt user to choose the roles to activate or deactivate
            Write-Host
            Write-Host "Enter the numbers corresponding to the roles you want to activate/deactivate (e.g., 1,3,5)"
            Write-Host "Or press 'R' to reload the PIM roles list or 'E' to return to the main menu."
            Write-Host "Press ESC at any time to return to the main menu."
            Write-Host -NoNewline "Choice: "
            try {
                $chosenIndicesInput = Read-UserInput -SpecialKeys @('E','R')
            } catch {
                if ($_.Exception.Message -eq 'UserCancelled') {
                    Write-Host "`nOperation cancelled by user. Returning to main menu." -ForegroundColor Yellow
                    Start-Sleep -Seconds 3
                    return
                } elseif ($_.Exception.Message -eq 'E') {
                    return
                } elseif ($_.Exception.Message -eq 'R') {
                    Write-Host "`nReloading PIM roles list..." -ForegroundColor Cyan
                    Start-Sleep -Seconds 1
                    continue  # Reload and display the roles list again
                } else {
                    Write-Host "`nAn unexpected error occurred: $_" -ForegroundColor Red
                    Start-Sleep -Seconds 3
                    return
                }
            }

            # Parse the input
            $chosenIndices = $chosenIndicesInput -split ',' | ForEach-Object { $_.Trim() }

            # Validate the selections
            $invalidSelections = $chosenIndices | Where-Object { -not ($_ -match '^\d+$') -or [int]$_ -lt 1 -or [int]$_ -gt $eligibleRoles.Count }
            if ($invalidSelections) {
                Write-Host "`nInvalid selection(s): $($invalidSelections -join ', ')" -ForegroundColor Red
                Start-Sleep -Seconds 2
                continue  # Re-prompt for selection
            }

            # Map selected indices to roles
            $selectedRoles = $chosenIndices | ForEach-Object { $roleIndexMap[[int]$_] }

            # Initialize lists for roles to activate and deactivate
            $rolesToActivate = @()
            $rolesToDeactivate = @()

            foreach ($selectedRole in $selectedRoles) {
                $selectedRoleId = $selectedRole.RoleDefinition.Id
                if ($activeRoleIds -contains $selectedRoleId) {
                    Write-Host "`nThe role '$($selectedRole.RoleDefinition.DisplayName)' is already active." -ForegroundColor Yellow
                    Write-Host "Do you want to deactivate this role? (Y/N)"
                    Write-Host -NoNewline "Choice: "
                    try {
                        $deactivateChoice = Read-UserInput
                        if ($deactivateChoice.ToUpper() -eq 'Y') {
                            $rolesToDeactivate += $selectedRole
                        } else {
                            Write-Host "Skipping deactivation of '$($selectedRole.RoleDefinition.DisplayName)'." -ForegroundColor Cyan
                        }
                    } catch {
                        if ($_.Exception.Message -eq 'UserCancelled') {
                            Write-Host "`nOperation cancelled by user." -ForegroundColor Yellow
                            Start-Sleep -Seconds 3
                            return
                        } else {
                            Write-Host "`nAn unexpected error occurred: $_" -ForegroundColor Red
                            Start-Sleep -Seconds 3
                            return
                        }
                    }
                } else {
                    $rolesToActivate += $selectedRole
                }
            }

            if (-not $rolesToActivate -and -not $rolesToDeactivate) {
                Write-Host "`nNo roles to activate or deactivate. Returning to role selection..." -ForegroundColor Cyan
                Start-Sleep -Seconds 2
                continue
            }

            # Proceed to deactivation for the selected roles
            if ($rolesToDeactivate) {
                foreach ($selectedRole in $rolesToDeactivate) {
                    try {
                        $selectedRoleId = $selectedRole.RoleDefinition.Id
                        # Find the active assignment schedule instance for the role
                        $activeAssignmentInstance = $activeRoles | Where-Object { $_.RoleDefinition.Id -eq $selectedRoleId }

                        if ($activeAssignmentInstance) {
                            # Deactivate the role by submitting a selfDeactivate request
                            $deactivateParams = @{
                                Action                = "selfDeactivate"
                                AssignmentScheduleId  = $activeAssignmentInstance.AssignmentScheduleId
                                Justification         = "User requested deactivation."
                                PrincipalId           = $currentUserId
                                RoleDefinitionId      = $selectedRoleId
                                DirectoryScopeId      = $selectedRole.DirectoryScopeId
                            }

                            # Submit the selfDeactivate request without ScheduleInfo or TicketInfo
                            New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $deactivateParams -ErrorAction Stop | Out-Null
                            Write-Host "Deactivation request for '$($selectedRole.RoleDefinition.DisplayName)' submitted successfully!" -ForegroundColor Green
                        } else {
                            Write-Host "Could not find the active assignment instance for this role." -ForegroundColor Red
                        }
                    } catch {
                        # Handle specific error: ActiveDurationTooShort
                        if ($_.Exception.Message -like '*ActiveDurationTooShort*') {
                            Write-Host "Cannot deactivate '$($selectedRole.RoleDefinition.DisplayName)'. The active duration is too short. Minimum required is 5 minutes." -ForegroundColor Red
                        } else {
                            Write-Host "An error occurred during deactivation of '$($selectedRole.RoleDefinition.DisplayName)':" -ForegroundColor Red
                            Write-Host $_.Exception.Message -ForegroundColor Red
                        }
                        # Continue with next role deactivation
                    }
                }

                # Reload the active roles after deactivation
                $activeRoles = Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance -Filter "principalId eq '$currentUserId'" -ExpandProperty RoleDefinition
                $activeRoleIds = $activeRoles | Select-Object -ExpandProperty RoleDefinition | Select-Object -ExpandProperty Id
            }

            # Proceed to activation for the selected roles
            if ($rolesToActivate) {
                # Prompt for Justification
                if ($rolesToActivate.Count -gt 1) {
                    Write-Host
                    Write-Host "Enter the justification for activating the selected roles (required):"
                } else {
                    Write-Host
                    Write-Host "Enter the justification for activating the role '$($rolesToActivate[0].RoleDefinition.DisplayName)' (required):"
                }
                Write-Host -NoNewline "Justification: "
                try {
                    $justification = Read-UserInput
                    while (-not $justification) {
                        Write-Host "Justification is required. Please enter a justification." -ForegroundColor Red
                        Write-Host -NoNewline "Justification: "
                        $justification = Read-UserInput
                    }
                } catch {
                    if ($_.Exception.Message -eq 'UserCancelled') {
                        Write-Host "`nOperation cancelled by user." -ForegroundColor Yellow
                        Start-Sleep -Seconds 3
                        return
                    } else {
                        Write-Host "`nAn unexpected error occurred: $_" -ForegroundColor Red
                        Start-Sleep -Seconds 3
                        return
                    }
                }

                # Check for TicketingRule (Assuming ticketing is required)
                $requiresTicket = $true  # Assume ticketing is required
                $ticketInfo = @{}

                # If ticketing is required, prompt for ticket details
                if ($requiresTicket) {
                    do {
                        Write-Host "Enter the ticket number (required):"
                        Write-Host -NoNewline "Ticket Number: "
                        try {
                            $ticketNumber = Read-UserInput
                            if (-not $ticketNumber) {
                                Write-Host "Ticket number is required. Please enter a valid ticket number." -ForegroundColor Red
                            }
                        } catch {
                            if ($_.Exception.Message -eq 'UserCancelled') {
                                Write-Host "`nOperation cancelled by user." -ForegroundColor Yellow
                                Start-Sleep -Seconds 3
                                return
                            } else {
                                Write-Host "`nAn unexpected error occurred: $_" -ForegroundColor Red
                                Start-Sleep -Seconds 3
                                return
                            }
                        }
                    } while (-not $ticketNumber)

                    Write-Host "Enter the ticket system (default is JIRA):"
                    Write-Host -NoNewline "Ticket System: "
                    try {
                        $ticketSystem = Read-UserInput
                        if (-not $ticketSystem) {
                            $ticketSystem = "JIRA"
                        }
                    } catch {
                        if ($_.Exception.Message -eq 'UserCancelled') {
                            Write-Host "`nOperation cancelled by user." -ForegroundColor Yellow
                            Start-Sleep -Seconds 3
                            return
                        } else {
                            Write-Host "`nAn unexpected error occurred: $_" -ForegroundColor Red
                            Start-Sleep -Seconds 3
                            return
                        }
                    }

                    $ticketInfo = @{
                        TicketNumber = $ticketNumber
                        TicketSystem = $ticketSystem
                    }
                }

                # Prompt for activation duration (default to 2 hours if none provided, with min=1 and max=8)
                do {
                    Write-Host "Enter the duration for the role activation in hours (1-8, default is 2 hours):"
                    Write-Host -NoNewline "Duration (hours): "
                    try {
                        $durationInput = Read-UserInput
                    } catch {
                        if ($_.Exception.Message -eq 'UserCancelled') {
                            Write-Host "`nOperation cancelled by user." -ForegroundColor Yellow
                            Start-Sleep -Seconds 3
                            return
                        } else {
                            Write-Host "`nAn unexpected error occurred: $_" -ForegroundColor Red
                            Start-Sleep -Seconds 3
                            return
                        }
                    }

                    if (-not $durationInput) {
                        $duration = "PT2H"  # Default to 2 hours
                        break
                    }

                    # Validate the duration input
                    if ($durationInput -match '^\d+$') {
                        $durationValue = [int]$durationInput
                        if ($durationValue -ge 1 -and $durationValue -le 8) {
                            $duration = "PT${durationValue}H"
                            break
                        } else {
                            Write-Host "Duration must be between 1 and 8 hours. Please try again." -ForegroundColor Red
                        }
                    } else {
                        Write-Host "Invalid input. Please enter a numeric value between 1 and 8." -ForegroundColor Red
                    }
                } while ($true)

                # Activate each role
                foreach ($selectedRole in $rolesToActivate) {
                    try {
                        $selectedRoleId = $selectedRole.RoleDefinition.Id

                        # Refresh active roles for accurate status
                        $activeRoles = Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance -Filter "principalId eq '$currentUserId'" -ExpandProperty RoleDefinition
                        $activeRoleIds = $activeRoles | Select-Object -ExpandProperty RoleDefinition | Select-Object -ExpandProperty Id

                        # Check if the role is now active
                        if ($activeRoleIds -contains $selectedRoleId) {
                            Write-Host "The role '$($selectedRole.RoleDefinition.DisplayName)' is already active." -ForegroundColor Yellow
                            continue
                        }

                        $activationParams = @{
                            Action            = "selfActivate"
                            PrincipalId       = $currentUserId
                            RoleDefinitionId  = $selectedRoleId
                            DirectoryScopeId  = $selectedRole.DirectoryScopeId
                            Justification     = $justification
                            ScheduleInfo      = @{
                                StartDateTime = (Get-Date)
                                Expiration    = @{
                                    Type     = "AfterDuration"
                                    Duration = $duration
                                }
                            }
                        }

                        # Add ticket info if required
                        if ($requiresTicket) {
                            $activationParams.TicketInfo = $ticketInfo
                        }

                        # Activate the role and suppress output
                        New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $activationParams -ErrorAction Stop | Out-Null
                        Write-Host "Activation request for '$($selectedRole.RoleDefinition.DisplayName)' submitted successfully!" -ForegroundColor Green

                    } catch {
                        # Check if error is due to existing role assignment
                        if ($_.Exception.Message -match 'RoleAssignmentExists') {
                            Write-Host "The role '$($selectedRole.RoleDefinition.DisplayName)' is already assigned or has a pending activation request." -ForegroundColor Yellow
                        } else {
                            Write-Host "An error occurred during activation of '$($selectedRole.RoleDefinition.DisplayName)':" -ForegroundColor Red
                            Write-Host $_.Exception.Message -ForegroundColor Red
                        }
                        # Continue with next role activation
                    }
                }

                # Reload the active roles after activation
                $activeRoles = Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance -Filter "principalId eq '$currentUserId'" -ExpandProperty RoleDefinition
                $activeRoleIds = $activeRoles | Select-Object -ExpandProperty RoleDefinition | Select-Object -ExpandProperty Id
            }

            # After actions, decide whether to continue, reload, or exit
            Write-Host
            Write-Host "Do you want to perform more actions? (Y/R/N)"
            Write-Host -NoNewline "Choice: "
            try {
                $continueChoice = Read-UserInput -SpecialKeys @('E','R','Y','N')
            } catch {
                if ($_.Exception.Message -eq 'UserCancelled') {
                    Write-Host "`nOperation cancelled by user." -ForegroundColor Yellow
                    Start-Sleep -Seconds 3
                    return
                } elseif ($_.Exception.Message -eq 'E') {
                    return
                } elseif ($_.Exception.Message -eq 'R' -or $_.Exception.Message -eq 'Y' -or $_.Exception.Message -eq 'N') {
                    $continueChoice = $_.Exception.Message
                } else {
                    Write-Host "`nAn unexpected error occurred: $_" -ForegroundColor Red
                    Start-Sleep -Seconds 3
                    return
                }
            }

            switch ($continueChoice.ToUpper()) {
                'Y' {
                    continue  # Loop back to reload and display the menu
                }
                'R' {
                    # Reload the PIM status by re-fetching roles
                    Write-Host "`nReloading PIM roles list..." -ForegroundColor Cyan
                    Start-Sleep -Seconds 1
                    continue  # Loop back to reload and display the updated list
                }
                'N' {
                    return  # Exit to the main menu
                }
                default {
                    Write-Host "Invalid choice. Exiting to the main menu." -ForegroundColor Red
                    return
                }
            }

        } catch {
            if ($_.Exception.Message -eq 'UserCancelled') {
                Write-Host "`nOperation cancelled by user." -ForegroundColor Yellow
                Start-Sleep -Seconds 3
                return
            } else {
                Write-Host "`nAn unexpected error occurred: $_" -ForegroundColor Red
                Start-Sleep -Seconds 3
                return
            }
        }
    }
}

# Search for users
function Search-Users {
    param ([string]$searchText)

    if (-not $global:isConnected) {
        Write-Host "Please connect first." -ForegroundColor Yellow
        Start-Sleep -Seconds 1
        return
    }

    $global:userResults = @()

    $escapedText = $searchText.Replace("'", "''")

    try {
        Write-Host "`nSearching users for '$searchText'..." -ForegroundColor Cyan
        $query = "https://graph.microsoft.com/v1.0/users?`$filter=startswith(userPrincipalName,'$escapedText') or startswith(displayName,'$escapedText') or startswith(givenName,'$escapedText') or startswith(surname,'$escapedText')"

        $morePages = $true
        while ($morePages) {
            $response = Invoke-MgGraphRequest -Method GET -Uri $query -ErrorAction Stop

            if ($response.value.Count -gt 0) {
                $global:userResults += $response.value
            }

            if ($response.'@odata.nextLink') {
                $query = $response.'@odata.nextLink'
            } else {
                $morePages = $false
            }
        }

        if ($global:userResults.Count -eq 1) {
            Write-Host "Found 1 user." -ForegroundColor Green
            Show-UserDetails -index 0
        } elseif ($global:userResults.Count -gt 1) {
            Write-Host "Found $($global:userResults.Count) users." -ForegroundColor Green
            Start-Sleep -Seconds 1
            Show-UserResults
        } else {
            Show-NotFoundMessage -Context 'Users'
            Start-Sleep -Seconds 1
        }
    } catch {
        Show-SimpleErrorMessage
        Start-Sleep -Seconds 1
    }
}

# Show User Results with Single-Key Navigation
function Show-UserResults {
    $global:totalPages = [math]::Ceiling($global:userResults.Count / $global:pageSize)
    $global:currentPage = 1

    while ($true) {
        Draw-Banner $global:userDisplayName
        Write-Host "Users - Page $global:currentPage/$global:totalPages`n" -ForegroundColor Cyan

        $start = ($global:currentPage -1) * $global:pageSize
        $end = [math]::Min($start + $global:pageSize, $global:userResults.Count)

        for ($i = $start; $i -lt $end; $i++) {
            $num = $i +1
            $user = $global:userResults[$i]
            Write-Host "$num. $($user.displayName), UPN: $($user.userPrincipalName)" -ForegroundColor White
        }

        Write-Host "`nOptions: [N] Next, [P] Previous, [V] View Details, [E] Main Menu" -ForegroundColor Yellow
        Write-Host "Press your choice..."
        $key = Read-SingleKey

        switch ($key) {
            'N' {
                if ($global:currentPage -lt $global:totalPages) {
                    $global:currentPage++
                } else {
                    Write-Host "Last page." -ForegroundColor Yellow
                    Start-Sleep -Seconds 1
                }
            }
            'P' {
                if ($global:currentPage -gt 1) {
                    $global:currentPage--
                } else {
                    Write-Host "First page." -ForegroundColor Yellow
                    Start-Sleep -Seconds 1
                }
            }
            'V' {
                Write-Host "`nEnter the number of the user to view details or press 'E' to return to Main Menu:"
                $selection = Read-Host "Choice"

                if ($selection.ToUpper() -eq 'E') {
                    return
                }

                if ($selection -match '^\d+$') {
                    $index = [int]$selection -1
                    if ($index -ge 0 -and $index -lt $global:userResults.Count) {
                        Show-UserDetails -index $index
                        # After viewing details, redraw the banner
                        Draw-Banner $global:userDisplayName
                    } else {
                        Write-Host "Invalid selection." -ForegroundColor Red
                        Start-Sleep -Seconds 1
                    }
                } else {
                    Write-Host "Invalid input." -ForegroundColor Red
                    Start-Sleep -Seconds 1
                }
            }
            'E' {
                return
            }
            'ENTER' {
                Write-Host "Invalid input." -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
            default {
                Write-Host "Invalid input." -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
    }
}

# Show User Details
function Show-UserDetails {
    param ([int]$index)

    try {
        $user = $global:userResults[$index]
        Clear-Host
        Write-Host ""
        Write-Host ("-" * 50) -ForegroundColor Yellow
        Write-Host "USER DETAILS" -ForegroundColor Cyan
        Write-Host ("-" * 50) -ForegroundColor Yellow
        Write-Host "DisplayName       : $($user.displayName)"
        Write-Host "GivenName         : $($user.givenName)"
        Write-Host "Surname           : $($user.surname)"
        Write-Host "JobTitle          : $($user.jobTitle)"
        Write-Host "Mail              : $($user.mail)"
        Write-Host "MobilePhone       : $($user.mobilePhone)"
        Write-Host "BusinessPhones    : $($($user.businessPhones -join ', '))"
        Write-Host "OfficeLocation    : $($user.officeLocation)"
        Write-Host "UserPrincipalName : $($user.userPrincipalName)"
        Write-Host ("-" * 50) -ForegroundColor Yellow

        # Fetch and display associated devices
        Show-UserDevices -userId $user.id

        # Prompt to return to the users list
        Write-Host "`nPress Enter to return to the users list." -ForegroundColor Yellow
        [void][System.Console]::ReadKey($true)
    } catch {
        Show-SimpleErrorMessage
        Write-Host "`nPress Enter to return to the users list." -ForegroundColor Yellow
        [void][System.Console]::ReadKey($true)
    }
}

# Show User Devices
function Show-UserDevices {
    param ([string]$userId)

    Write-Host ""
    Write-Host ("-" * 50) -ForegroundColor Yellow
    Write-Host "REGISTERED DEVICES" -ForegroundColor Cyan
    Write-Host ("-" * 50) -ForegroundColor Yellow

    try {
        $query = "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($userId))/managedDevices"

        $morePages = $true
        while ($morePages) {
            $response = Invoke-MgGraphRequest -Method GET -Uri $query -ErrorAction Stop

            if ($response -and $response.value -and $response.value.Count -gt 0) {
                foreach ($device in $response.value) {
                    Write-Host "Device Name   : $($device.deviceName)"
                    Write-Host "Model         : $($device.model)"
                    Write-Host "Serial Number : $($device.serialNumber)"
                    Write-Host ("-" * 50) -ForegroundColor Yellow
                }
            }

            if ($response.'@odata.nextLink') {
                $query = $response.'@odata.nextLink'
            } else {
                $morePages = $false
            }
        }

        if (-not $response.value -or $response.value.Count -eq 0) {
            Show-NotFoundMessage -Context 'Devices'
        }
    } catch {
        Show-SimpleErrorMessage
    }
}

# Search Devices
function Search-Devices {
    param ([string]$searchText)

    if (-not $global:isConnected) {
        Write-Host "Please connect first." -ForegroundColor Yellow
        Start-Sleep -Seconds 1
        return
    }

    $global:deviceResults = @()

    $escapedText = $searchText.Replace("'", "''")

    try {
        Write-Host "`nSearching devices for '$searchText'..." -ForegroundColor Cyan

        # First, search by serial number
        $querySerial = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=serialNumber eq '$escapedText'"

        $morePages = $true
        while ($morePages) {
            $responseSerial = Invoke-MgGraphRequest -Method GET -Uri $querySerial -ErrorAction Stop

            if ($responseSerial.value.Count -gt 0) {
                $global:deviceResults += $responseSerial.value
            }

            if ($responseSerial.'@odata.nextLink') {
                $querySerial = $responseSerial.'@odata.nextLink'
            } else {
                $morePages = $false
            }
        }

        if ($global:deviceResults.Count -eq 1) {
            Write-Host "Found 1 device by Serial Number." -ForegroundColor Green
            Show-DeviceDetails -index 0
            return
        } elseif ($global:deviceResults.Count -gt 1) {
            Write-Host "Found $($global:deviceResults.Count) devices by Serial Number." -ForegroundColor Green
            Start-Sleep -Seconds 1
            Show-DeviceResults
            return
        } else {
            Write-Host "No devices found by Serial Number." -ForegroundColor Yellow
            # Clear deviceResults for next search
            $global:deviceResults = @()
            # Now, search by device name
            $queryName = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=startswith(deviceName,'$escapedText')"

            $morePages = $true
            while ($morePages) {
                $responseName = Invoke-MgGraphRequest -Method GET -Uri $queryName -ErrorAction Stop

                if ($responseName.value.Count -gt 0) {
                    $global:deviceResults += $responseName.value
                }

                if ($responseName.'@odata.nextLink') {
                    $queryName = $responseName.'@odata.nextLink'
                } else {
                    $morePages = $false
                }
            }

            if ($global:deviceResults.Count -eq 1) {
                Write-Host "Found 1 device by Name." -ForegroundColor Green
                Show-DeviceDetails -index 0
            } elseif ($global:deviceResults.Count -gt 1) {
                Write-Host "Found $($global:deviceResults.Count) devices by Name." -ForegroundColor Green
                Start-Sleep -Seconds 1
                Show-DeviceResults
            } else {
                Show-NotFoundMessage -Context 'Devices'
                Start-Sleep -Seconds 1
            }
        }

    } catch {
        Show-SimpleErrorMessage
        Start-Sleep -Seconds 1
    }
}

# Show Device Results with Single-Key Navigation
function Show-DeviceResults {
    $global:totalPages = [math]::Ceiling($global:deviceResults.Count / $global:pageSize)
    $global:currentPage = 1

    while ($true) {
        Draw-Banner $global:userDisplayName
        Write-Host "Devices - Page $global:currentPage/$global:totalPages`n" -ForegroundColor Cyan

        $start = ($global:currentPage -1) * $global:pageSize
        $end = [math]::Min($start + $global:pageSize, $global:deviceResults.Count)

        for ($i = $start; $i -lt $end; $i++) {
            $num = $i +1
            $device = $global:deviceResults[$i]
            Write-Host "$num. $($device.deviceName), Serial: $($device.serialNumber)" -ForegroundColor White
        }

        Write-Host "`nOptions: [N] Next, [P] Previous, [V] View Details, [E] Main Menu" -ForegroundColor Yellow
        Write-Host "Press your choice..."
        $key = Read-SingleKey

        switch ($key) {
            'N' {
                if ($global:currentPage -lt $global:totalPages) {
                    $global:currentPage++
                } else {
                    Write-Host "Last page." -ForegroundColor Yellow
                    Start-Sleep -Seconds 1
                }
            }
            'P' {
                if ($global:currentPage -gt 1) {
                    $global:currentPage--
                } else {
                    Write-Host "First page." -ForegroundColor Yellow
                    Start-Sleep -Seconds 1
                }
            }
            'V' {
                Write-Host "`nEnter the number of the device to view details or press 'E' to return to Main Menu:"
                $selection = Read-Host "Choice"

                if ($selection.ToUpper() -eq 'E') {
                    return
                }

                if ($selection -match '^\d+$') {
                    $index = [int]$selection -1
                    if ($index -ge 0 -and $index -lt $global:deviceResults.Count) {
                        Show-DeviceDetails -index $index
                        # After viewing details, redraw the banner
                        Draw-Banner $global:userDisplayName
                    } else {
                        Write-Host "Invalid selection." -ForegroundColor Red
                        Start-Sleep -Seconds 1
                    }
                } else {
                    Write-Host "Invalid input." -ForegroundColor Red
                    Start-Sleep -Seconds 1
                }
            }
            'E' {
                return
            }
            'ENTER' {
                Write-Host "Invalid input." -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
            default {
                Write-Host "Invalid input." -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
    }
}

# Show Device Details
function Show-DeviceDetails {
    param ([int]$index)

    try {
        $device = $global:deviceResults[$index]
        Clear-Host
        Write-Host ""
        Write-Host ("-" * 50) -ForegroundColor Yellow
        Write-Host "DEVICE DETAILS" -ForegroundColor Cyan
        Write-Host ("-" * 50) -ForegroundColor Yellow
        Write-Host "Device Name   : $($device.deviceName)"
        Write-Host "Model         : $($device.model)"
        Write-Host "Serial Number : $($device.serialNumber)"
        Write-Host ("-" * 50) -ForegroundColor Yellow

        # Fetch and display associated user
        Show-DeviceUser -userId $device.userId

        # Prompt to return to the devices list
        Write-Host "`nPress Enter to return to the devices list." -ForegroundColor Yellow
        [void][System.Console]::ReadKey($true)
    } catch {
        Show-SimpleErrorMessage
        Write-Host "`nPress Enter to return to the devices list." -ForegroundColor Yellow
        [void][System.Console]::ReadKey($true)
    }
}

# Show Device User
function Show-DeviceUser {
    param ([string]$userId)

    Write-Host ""
    Write-Host ("-" * 50) -ForegroundColor Yellow
    Write-Host "ASSOCIATED USER" -ForegroundColor Cyan
    Write-Host ("-" * 50) -ForegroundColor Yellow

    if ([string]::IsNullOrEmpty($userId)) {
        Write-Host "No user associated." -ForegroundColor Yellow
        return
    }

    try {
        $query = "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($userId))"
        $response = Invoke-MgGraphRequest -Method GET -Uri $query -ErrorAction Stop

        Write-Host "UserPrincipalName : $($response.userPrincipalName)"
        Write-Host "DisplayName       : $($response.displayName)"
        Write-Host ("-" * 50) -ForegroundColor Yellow
    } catch {
        Show-SimpleErrorMessage
    }
}

# Disconnect
function Disconnect-Graph {
    Write-Host "`nDisconnecting..." -ForegroundColor Cyan
    Disconnect-MgGraph
    $global:isConnected = $false
    $global:userDisplayName = ""
    Write-Host "Disconnected." -ForegroundColor Green
}

# Main Loop
do {
    Draw-Banner
    Write-Host "Press any key to authenticate or 'E' to exit." -ForegroundColor Yellow

    $startKey = Read-SingleKey

    if ($startKey -eq 'E') {
        if ($global:isConnected) {
            Disconnect-Graph
        }
        Write-Host "Goodbye!" -ForegroundColor Green
        exit
    } else {
        Authenticate-Graph
    }

    if ($global:isConnected) {
        while ($global:isConnected) {
            Draw-Banner $global:userDisplayName
            Write-Host "----------------------------------------" -ForegroundColor Yellow
            Write-Host "1. [P]IM Management"
            Write-Host "2. Search [U]sers"
            Write-Host "3. Search [D]evices"
            Write-Host "4. [E]xit"

            Write-Host "`nOptions: [1] Activate PIM, [2] Search Users, [3] Search Devices, [4] Exit"
            Write-Host "Enter selection:"
            $choice = Read-SingleKey

            switch ($choice.ToUpper()) {
                '1' {
                    Activate-PIM
                    # After activation, redraw the banner
                    Draw-Banner $global:userDisplayName
                }
                'P' {
                    Activate-PIM
                    # After activation, redraw the banner
                    Draw-Banner $global:userDisplayName
                }
                '2' {
                    Write-Host "`nEnter search text or press 'E' to return to Main Menu:"
                    $searchText = Read-Host "Search UPN/Name"

                    if ($searchText.ToUpper() -eq 'E') {
                        continue
                    }
                    if (-not [string]::IsNullOrWhiteSpace($searchText)) {
                        Search-Users -searchText $searchText
                        # After search, redraw the banner
                        Draw-Banner $global:userDisplayName
                    } else {
                        Write-Host "Enter valid search text." -ForegroundColor Yellow
                        Start-Sleep -Seconds 1
                    }
                }
                'U' {
                    Write-Host "`nEnter search text or press 'E' to return to Main Menu:"
                    $searchText = Read-Host "Search UPN/Name"

                    if ($searchText.ToUpper() -eq 'E') {
                        continue
                    }
                    if (-not [string]::IsNullOrWhiteSpace($searchText)) {
                        Search-Users -searchText $searchText
                        # After search, redraw the banner
                        Draw-Banner $global:userDisplayName
                    } else {
                        Write-Host "Enter valid search text." -ForegroundColor Yellow
                        Start-Sleep -Seconds 1
                    }
                }
                '3' {
                    Write-Host "`nEnter search text or press 'E' to return to Main Menu:"
                    $searchText = Read-Host "Search Device/Serial"

                    if ($searchText.ToUpper() -eq 'E') {
                        continue
                    }
                    if (-not [string]::IsNullOrWhiteSpace($searchText)) {
                        Search-Devices -searchText $searchText
                        # After search, redraw the banner
                        Draw-Banner $global:userDisplayName
                    } else {
                        Write-Host "Enter valid search text." -ForegroundColor Yellow
                        Start-Sleep -Seconds 1
                    }
                }
                'D' {
                    Write-Host "`nEnter search text or press 'E' to return to Main Menu:"
                    $searchText = Read-Host "Search Device/Serial"

                    if ($searchText.ToUpper() -eq 'E') {
                        continue
                    }
                    if (-not [string]::IsNullOrWhiteSpace($searchText)) {
                        Search-Devices -searchText $searchText
                        # After search, redraw the banner
                        Draw-Banner $global:userDisplayName
                    } else {
                        Write-Host "Enter valid search text." -ForegroundColor Yellow
                        Start-Sleep -Seconds 1
                    }
                }
                '4' {
                    if ($global:isConnected) {
                        Disconnect-Graph | Out-Null
                    }
                    Write-Host "Goodbye!" -ForegroundColor Green
                    exit
                }
                'E' {
                    if ($global:isConnected) {
                        Disconnect-Graph | Out-Null
                    }
                    Write-Host "Goodbye!" -ForegroundColor Green
                    exit
                }
                'ENTER' {
                    Write-Host "Invalid choice." -ForegroundColor Red
                    Start-Sleep -Seconds 1
                }
                default {
                    Write-Host "Invalid choice." -ForegroundColor Red
                    Start-Sleep -Seconds 1
                }
            }
        }
    } else {
        Write-Host "Authentication failed. Try again." -ForegroundColor Red
        Write-Host "Press Enter to retry or 'E' to exit." -ForegroundColor Yellow
        $retryKey = Read-SingleKey

        if ($retryKey -eq 'E') {
            Write-Host "Goodbye!" -ForegroundColor Green
            exit
        }
    }
} while ($true)
