# Required Modules
# Automatically install specific Microsoft.Graph modules if not installed
$requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.DeviceManagement")

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
$graphScopes = @("User.Read.All", "DeviceManagementManagedDevices.Read.All")

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

# Function to Draw a Consistent ASCII Art Banner
function Draw-Banner {
    param (
        [string]$ConnectedUser = ""
    )

    Clear-Host

    # ASCII Art Banner
    $banner = @"
  ____                 _       ____            _             
 |  _ \               | |     |  _ \          | |            
 | |_) | ___  ___  ___| |_ ___| |_) | ___  ___| | _____ _ __ 
 |  _ < / _ \/ __|/ _ \ __/ _ \  _ < / _ \/ __| |/ / _ \ '__|
 | |_) | (_) \__ \  __/ ||  __/ |_) |  __/ (__|   <  __/ |   
 |____/ \___/|___/\___|\__\___|____/ \___|\___|_|\_\___|_|   
"@

    Write-Host $banner -ForegroundColor Green
    Write-Host ("=" * ($banner.Split("`n")[0].Length)) -ForegroundColor Green

    if (-not [string]::IsNullOrEmpty($ConnectedUser)) {
        Write-Host "`nConnected as: $ConnectedUser" -ForegroundColor Cyan
    }

    Write-Host "`n"
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
            $global:userDisplayName = $context.Account.Username
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

        # Draw a box around the user list using ASCII characters
        $boxWidth = 60
        $topBorder = "+" + ("-" * ($boxWidth -2)) + "+"
        $bottomBorder = "+" + ("-" * ($boxWidth -2)) + "+"

        Write-Host $topBorder -ForegroundColor Green

        for ($i = $start; $i -lt $end; $i++) {
            $num = $i +1
            $user = $global:userResults[$i]
            $line = "$num. $($user.displayName), UPN: $($user.userPrincipalName)"
            # Truncate the line if it's longer than boxWidth -4 to fit within the box
            if ($line.Length -gt ($boxWidth -4)) {
                $line = $line.Substring(0, $boxWidth -7) + "..."
            }
            # Pad the line to fit the box
            $paddedLine = $line.PadRight($boxWidth -4)
            Write-Host "| $paddedLine |" -ForegroundColor White
        }

        # If fewer items on the last page, fill the rest with empty lines
        for ($j = $end; $j -lt ($start + $global:pageSize); $j++) {
            Write-Host "| " + (" " * ($boxWidth -4)) + " |" -ForegroundColor White
        }

        Write-Host $bottomBorder -ForegroundColor Green

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
        Write-Host "`nUser Details:" -ForegroundColor Cyan
        Write-Host "+----------------------------------------+" -ForegroundColor Green
        Write-Host "| DisplayName       : $($user.displayName.PadRight(32)) |" -ForegroundColor White
        Write-Host "| GivenName         : $($user.givenName.PadRight(32)) |" -ForegroundColor White
        Write-Host "| Surname           : $($user.surname.PadRight(32)) |" -ForegroundColor White
        Write-Host "| JobTitle          : $($user.jobTitle.PadRight(32)) |" -ForegroundColor White
        Write-Host "| Mail              : $($user.mail.PadRight(32)) |" -ForegroundColor White
        Write-Host "| MobilePhone       : $($user.mobilePhone.PadRight(32)) |" -ForegroundColor White
        Write-Host "| BusinessPhones    : $($user.businessPhones -join ', '.PadRight(32)) |" -ForegroundColor White
        Write-Host "| OfficeLocation    : $($user.officeLocation.PadRight(32)) |" -ForegroundColor White
        Write-Host "| UserPrincipalName : $($user.userPrincipalName.PadRight(24)) |" -ForegroundColor White
        Write-Host "+----------------------------------------+" -ForegroundColor Green

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

    Write-Host "`nRegistered Devices:" -ForegroundColor Cyan

    try {
        $query = "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($userId))/managedDevices"

        $morePages = $true
        $devices = @()
        while ($morePages) {
            $response = Invoke-MgGraphRequest -Method GET -Uri $query -ErrorAction Stop

            if ($response -and $response.value -and $response.value.Count -gt 0) {
                $devices += $response.value
            }

            if ($response.'@odata.nextLink') {
                $query = $response.'@odata.nextLink'
            } else {
                $morePages = $false
            }
        }

        if ($devices.Count -gt 0) {
            foreach ($device in $devices) {
                Write-Host "+------------------------------+" -ForegroundColor Green
                Write-Host "| Device Name   : $($device.deviceName.PadRight(24)) |" -ForegroundColor White
                Write-Host "| Model         : $($device.model.PadRight(24)) |" -ForegroundColor White
                Write-Host "| Serial Number : $($device.serialNumber.PadRight(24)) |" -ForegroundColor White
                Write-Host "+------------------------------+" -ForegroundColor Green
            }
        } else {
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

        # Draw a box around the device list using ASCII characters
        $boxWidth = 60
        $topBorder = "+" + ("-" * ($boxWidth -2)) + "+"
        $bottomBorder = "+" + ("-" * ($boxWidth -2)) + "+"

        Write-Host $topBorder -ForegroundColor Green

        for ($i = $start; $i -lt $end; $i++) {
            $num = $i +1
            $device = $global:deviceResults[$i]
            $line = "$num. $($device.deviceName), Serial: $($device.serialNumber)"
            # Truncate the line if it's longer than boxWidth -4 to fit within the box
            if ($line.Length -gt ($boxWidth -4)) {
                $line = $line.Substring(0, $boxWidth -7) + "..."
            }
            # Pad the line to fit the box
            $paddedLine = $line.PadRight($boxWidth -4)
            Write-Host "| $paddedLine |" -ForegroundColor White
        }

        # If fewer items on the last page, fill the rest with empty lines
        for ($j = $end; $j -lt ($start + $global:pageSize); $j++) {
            Write-Host "| " + (" " * ($boxWidth -4)) + " |" -ForegroundColor White
        }

        Write-Host $bottomBorder -ForegroundColor Green

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
        Write-Host "`nDevice Details:" -ForegroundColor Cyan
        Write-Host "+------------------------------+" -ForegroundColor Green
        Write-Host "| Device Name   : $($device.deviceName.PadRight(24)) |" -ForegroundColor White
        Write-Host "| Model         : $($device.model.PadRight(24)) |" -ForegroundColor White
        Write-Host "| Serial Number : $($device.serialNumber.PadRight(24)) |" -ForegroundColor White
        Write-Host "+------------------------------+" -ForegroundColor Green

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

    Write-Host "`nAssociated User:" -ForegroundColor Cyan

    if ([string]::IsNullOrEmpty($userId)) {
        Write-Host "No user associated." -ForegroundColor Yellow
        return
    }

    try {
        $query = "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($userId))"
        $response = Invoke-MgGraphRequest -Method GET -Uri $query -ErrorAction Stop

        Write-Host "| UserPrincipalName : $($response.userPrincipalName.PadRight(24)) |" -ForegroundColor White
        Write-Host "| DisplayName       : $($response.displayName.PadRight(24)) |" -ForegroundColor White
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
            Write-Host "+----------------------------------------------+" -ForegroundColor Green
            Write-Host "| 1. Search [U]sers                            |" -ForegroundColor White
            Write-Host "| 2. Search [D]evices                          |" -ForegroundColor White
            Write-Host "| 3. [E]xit                                    |" -ForegroundColor White
            Write-Host "+----------------------------------------------+" -ForegroundColor Green

            Write-Host "`nOptions: [1] Search Users, [2] Search Devices, [3] Exit" -ForegroundColor Yellow
            Write-Host "Enter selection:"
            $choice = Read-SingleKey

            switch ($choice.ToUpper()) {
                '1' {
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
                '2' {
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
                '3' {
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
