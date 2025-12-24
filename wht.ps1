<#
.SYNOPSIS
    Hotspot Toggle with Persistent WiFi Adapter Selection
.DESCRIPTION
    Remembers selected WiFi adapter between sessions only when manual selection is needed
.NOTES
    Added BurntToast notification support for non-interactive execution.
#>

# ==== Assembly Loading (Fixes TypeNotFound Error) ====
try {
    # Load the necessary assembly for WinRT / AsTask support immediately at runtime
    Add-Type -AssemblyName "System.Runtime.WindowsRuntime" -ErrorAction Stop
} catch {
    [System.Reflection.Assembly]::Load("System.Runtime.WindowsRuntime, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089") | Out-Null
}

# Load WinRT namespaces for Radio and Networking
[Windows.Networking.Connectivity.NetworkInformation, Windows.Networking.Connectivity, ContentType = WindowsRuntime] | Out-Null
[Windows.Networking.NetworkOperators.NetworkOperatorTetheringManager, Windows.Networking.NetworkOperators, ContentType = WindowsRuntime] | Out-Null
[Windows.Devices.Radios.Radio, Windows.System.Devices, ContentType = WindowsRuntime] | Out-Null
[Windows.Devices.Radios.RadioAccessStatus, Windows.System.Devices, ContentType = WindowsRuntime] | Out-Null
[Windows.Devices.Radios.RadioState, Windows.System.Devices, ContentType = WindowsRuntime] | Out-Null

# ==== Script Version ====

    # This is automatically updated via pre-commit hook
    $scriptVersion = "1.0.8"

    # Config file path
    $configFile = "$PSScriptRoot\adapter.config"

    # Create log file? For debugging purposes
    $logFile = $true

    # should the log file be printed in reverse order?
    $logReverse = $false

    # Define a log file path
    $logFilePath = "$PSScriptRoot\script.log"

    # Restart Wifi adapter? 
    # This is necessary for certain firewalls such as Comodo to open ports for the Hotspot
    $restartWiFi = $true
    
    # Force Administrator mode
    $forceAsAdmin = $true
    
    # Add delay in Task scheduler/startup mode
    $delay = 30

# ==== WinRT Helper Logic ====

    # Helper to await WinRT tasks in PowerShell
    # We fetch the type via Reflection to avoid "TypeNotFound" parsing errors
    $RuntimeExtensions = [AppDomain]::CurrentDomain.GetAssemblies() | 
        ForEach-Object { $_.GetType("System.WindowsRuntimeSystemExtensions") } | 
        Where-Object { $_ -ne $null } | Select-Object -First 1

    if ($RuntimeExtensions) {
        $asTaskGeneric = $RuntimeExtensions.GetMethods() | Where-Object {
            $_.Name -eq 'AsTask' -and
            $_.GetParameters().Count -eq 1 -and
            $_.GetParameters()[0].ParameterType.Name -like 'IAsyncOperation*'
        } | Select-Object -First 1
    }

    function Await-WinRT {
        param (
            $winRtOperation,
            [Type]$resultType
        )

        try {
            if (-not $asTaskGeneric) {
                Write-Warning "Await-WinRT: Failed to find AsTask method. Ensure System.Runtime.WindowsRuntime is loaded."
                return $null
            }

            $asTask = $asTaskGeneric.MakeGenericMethod($resultType)
            $netTask = $asTask.Invoke($null, @($winRtOperation))
            $netTask.Wait(-1) | Out-Null
            return $netTask.Result
        } catch {
            Write-Warning "Await-WinRT failed: $_"
            return $null
        }
    }

# ==== FUNCTIONS =====

    # Notification Function
    function Send-Notification {
        param (
            [string]$Title,
            [string]$Message
        )
        
        # Only attempt if BurntToast module is available
        if (Get-Module -ListAvailable -Name BurntToast) {
            New-BurntToastNotification -Text $Title, $Message
        } else {
            LogThis "BurntToast module not found. Notification skipped." -Color "Yellow"
        }
    }

    # Function to check if the script is running as admin
    function IsAdmin {

        $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }

    # Run script as Administrator
    function RunAsAdmin {

        # Skip elevation if running as SYSTEM user
        if ($env:SYSTEMROOT -and $env:USERNAME -eq "SYSTEM") {
            LogThis "Running as SYSTEM via Task Scheduler. Skipping elevation check." -verboseMessage $true
            return
        }

        # Relaunch script as admin if not already running as admin
        if (-Not (IsAdmin)) {

            Write-Host "This script requires administrative privileges. Requesting elevation..." -ForegroundColor Yellow
            Start-Process -FilePath "powershell.exe" `
                          -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" `
                          -Verb RunAs
            Exit 0
        }
    }

    # Create the logging system
    function LogThis  {
        param (
            [string]$Message,
            [string]$Color = "White"
        )
        
        #Terminal message
        Write-Host $Message -ForegroundColor $Color
        
        # Log only if Logging is enabled
        if ($logFile) {
            #check if printing in reverse or not
            if ($logReverse) {
                
                # Read the existing content of the log file
                $existingContent = Get-Content -Path $logFilePath -Raw

                # Prepend the new log entry to the existing content
                $updatedContent = "$Message`n$existingContent"

                # Write the updated content back to the log file
                $updatedContent | Set-Content -Path $logFilePath -Encoding UTF8
                
            } else {
                
                $Message | Out-File -Append -FilePath $logFilePath -Encoding UTF8
            }

        }
    }

    # Determine if the script runs interactively
    function IsRunningFromTerminal {

        # Get the current process ID
        $proc = Get-CimInstance Win32_Process -Filter "ProcessId = $pid"

        # Get the parent process ID
        $parent = Get-CimInstance Win32_Process -Filter "ProcessId = $($proc.ParentProcessId)"

        # Check if the parent process is 'svchost.exe' and contains 'Schedule' in the command line
        if ($parent.Name -eq "svchost.exe" -or $parent.CommandLine -like "*Schedule*") {

            return $false

        } else {

            return $true
        }
    }
    
    # Identify Wifi Adapter
    function Get-WiFiAdapter {
        # Try to load saved adapter first (only if config exists)
        if (Test-Path $configFile) {
            try {
                LogThis "Found config file, attempting to load adapter"                                                            
                $savedAdapter = Get-Content $configFile -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                $adapter = Get-NetAdapter | Where-Object { 
                    $_.Name -eq $savedAdapter.Name -and 
                    $_.InterfaceDescription -eq $savedAdapter.Description
                }
                
                if ($adapter) {
                    LogThis -Message "Using saved WiFi adapter: $($adapter.Name)"                                                      
                    return $adapter
                } else {
                    LogThis "Saved adapter not found"
                }
            }
            catch {
                LogThis "Error reading adapter config: $_"  -Color "Yellow"
           }
        }

        # Automatic detection patterns
        $patterns = @(
            "*Wireless*", 
            "*Wi-Fi*",
            "*WLAN*",
            "*802.11*"
        )
        
        foreach ($pattern in $patterns) {
            $adapter = Get-NetAdapter | Where-Object { 
                $_.Name -like $pattern -or 
                $_.InterfaceDescription -like $pattern
            } | Select-Object -First 1
            
            if ($adapter) {
                LogThis "Found WiFi adapter via pattern: $pattern -> $($adapter.Name)"
                return $adapter
            }
        }

        # Fallback to physical media types
        $adapter = Get-NetAdapter | Where-Object {
            $_.PhysicalMediaType -eq 'Native 802.11' -or
            $_.MediaType -eq 'Wireless WAN'
        } | Select-Object -First 1
        
        if ($adapter) {
            LogThis "Found WiFi adapter via fallback type match: $($adapter.Name)"                                                         
            return $adapter
        }

        # Final fallback - manual selection (only creates config in this case)
        LogThis -Message "Automatic detection failed, requesting manual adapter selection" -Color "Yellow"
        $selected = Get-NetAdapter | Out-GridView -Title "Select your WiFi adapter (Will be remembered for future use)" -PassThru
        
        if ($selected) {
            LogThis "Manual adapter selected: $($selected.Name)"
            Save-AdapterConfig $selected
            return $selected
        }

        LogThis "No WiFi adapter selected, aborting"  -Color "Red"
        throw "No WiFi adapter selected"
    }

    # Save WIFI adapter in a file, if manually selected
    function Save-AdapterConfig {
        param($adapter)
        
        try {
            if (-not (Test-Path $PSScriptRoot)) {
                New-Item -ItemType Directory -Path $PSScriptRoot -ErrorAction Stop | Out-Null
            }
            
            [PSCustomObject]@{
                Name = $adapter.Name
                Description = $adapter.InterfaceDescription
            } | ConvertTo-Json -ErrorAction Stop | Out-File $configFile -ErrorAction Stop
            
            LogThis "Adapter config saved: $($adapter.Name)" -Color "Green"
        }
        catch {
            LogThis "Failed to save adapter config: $_" -Color "Yellow"
        }
    }
    
    # Restart WiFi Adapter (Necessary for certain firewalls such as Comodo)
    function Restart-WiFi {
        try {
            $wifiAdapter = Get-NetAdapter | Where-Object { $_.InterfaceDescription -like "*Wireless*" } | Select-Object -First 1
            if (-not $wifiAdapter) {
                Write-Host "WiFi adapter not found"  -Color "Yellow"
                return $false
            }

            LogThis "Restarting WiFi adapter..."  -Color "Yellow"
            Restart-NetAdapter -Name $wifiAdapter.Name -Confirm:$false -ErrorAction Stop
            Start-Sleep -Seconds 2
            LogThis "WiFi restart complete"  -Color "Green"
            return $true
        }
        catch {
            LogThis "WiFi restart failed: $_" -Color "Red"
            return $false
        }
    }

    # Main function
    function old-Toggle-Hotspot {
        try {
            $connectionProfile = [Windows.Networking.Connectivity.NetworkInformation]::GetInternetConnectionProfile()
            $tetheringManager = [Windows.Networking.NetworkOperators.NetworkOperatorTetheringManager]::CreateFromConnectionProfile($connectionProfile)
            
            if ($tetheringManager.TetheringOperationalState -eq 'On') {
                LogThis "Hotspot is active - disabling..." -Color "Yellow"
                $null = $tetheringManager.StopTetheringAsync()
                LogThis "Hotspot disabled"  -Color "Green"
                Send-Notification -Title "Hotspot Disabled" -Message "The mobile hotspot has been turned off."
            }
            else {
                LogThis "Hotspot is inactive - enabling..."  -Color "Yellow"
                if ($restartWiFi) {Restart-WiFi | Out-Null}
                $null = $tetheringManager.StartTetheringAsync()
                LogThis "Hotspot enabled"  -Color "Green"
                Send-Notification -Title "Hotspot Enabled" -Message "The mobile hotspot is now active."
            }
            return $true
        }
        catch {
            LogThis "Hotspot toggle failed: $_" -Color "Red"
            return $false
        }
    }

    function Toggle-Hotspot {
        try {
            # (1) Create / retrieve the tethering manager for the current Internet connection
            $connectionProfile = [Windows.Networking.Connectivity.NetworkInformation]::GetInternetConnectionProfile()
            $tetheringManager   = [Windows.Networking.NetworkOperators.NetworkOperatorTetheringManager]::CreateFromConnectionProfile($connectionProfile)

            # (2) If hotspot is already ON, just turn it OFF and exit
            if ($tetheringManager.TetheringOperationalState -eq 'On') {
                LogThis "Hotspot is active -> disabling..." -Color "Yellow"
                $null = Await-WinRT ($tetheringManager.StopTetheringAsync()) ([Windows.Networking.NetworkOperators.NetworkOperatorTetheringOperationResult])
                LogThis "Hotspot disabled" -Color "Green"
                Send-Notification -Title "Hotspot Disabled" -Message "The mobile hotspot has been turned off."
                return $true
            }

            # From here onward, Hotspot is currently OFF. We want to enable it, but first ensure
            #    the Wi-Fi radio is in the “enabled but OFF” state. If it’s already OFF, we can skip.

            # (5) Request access (usually granted) and enumerate all radios
            Await-WinRT ([Windows.Devices.Radios.Radio]::RequestAccessAsync()) ([Windows.Devices.Radios.RadioAccessStatus]) | Out-Null
            $allRadios = Await-WinRT ([Windows.Devices.Radios.Radio]::GetRadiosAsync()) ([System.Collections.Generic.IReadOnlyList[Windows.Devices.Radios.Radio]])

            # (6) Find the Wi-Fi radio object
            $wifiRadio = $allRadios | Where-Object { $_.Kind -eq 'WiFi' }
            if (-not $wifiRadio) {
                LogThis "ERROR: No Wi-Fi radio found via WinRT. Cannot force radio OFF." -Color "Red"
                return $false
            }

            # (7) If Wi-Fi radio is already OFF, skip the hack. Otherwise, do the “start & stop hotspot” dance.
            if ($wifiRadio.State -eq 'On') {
                LogThis "Wi-Fi radio is currently ON -> forcing ‘radio OFF’ via Hotspot hack..." -Color "Yellow"

                # (7a) Temporary: start Hotspot. That will force the Wi-Fi radio ON (if not already).
                $null = Await-WinRT ($tetheringManager.StartTetheringAsync()) ([Windows.Networking.NetworkOperators.NetworkOperatorTetheringOperationResult])
                LogThis "  -> temporary hotspot started" -Color "Cyan"
                Start-Sleep -Seconds 2

                # (7b) Immediately stop Hotspot. Windows will leave the Wi-Fi radio in “enabled but OFF” state.
                $null = Await-WinRT ($tetheringManager.StopTetheringAsync()) ([Windows.Networking.NetworkOperators.NetworkOperatorTetheringOperationResult])
                LogThis "  -> temporary hotspot stopped; Wi-Fi radio should now be OFF" -Color "Cyan"
                Start-Sleep -Seconds 2

                # (7c) Re‐check the Wi-Fi radio state
                $allRadios = Await-WinRT ([Windows.Devices.Radios.Radio]::GetRadiosAsync()) ([System.Collections.Generic.IReadOnlyList[Windows.Devices.Radios.Radio]])
                $wifiRadio = $allRadios | Where-Object { $_.Kind -eq 'WiFi' }
                if ($wifiRadio.State -ne 'Off') {
                    LogThis "Failed to switch Wi-Fi radio to OFF. Aborting hotspot start." -Color "Red"
                    return $false
                }
                LogThis "Wi-Fi radio is now OFF (but adapter still enabled)" -Color "Green"
            }
            else {
                LogThis "Wi-Fi radio is already OFF; proceeding to start Hotspot." -Color "Cyan"
            }

            # (8) Finally, start the Hotspot for real
            LogThis "Hotspot is inactive -> enabling now..." -Color "Yellow"
            $null = Await-WinRT ($tetheringManager.StartTetheringAsync()) ([Windows.Networking.NetworkOperators.NetworkOperatorTetheringOperationResult])
            LogThis "Hotspot enabled" -Color "Green"
            Send-Notification -Title "Hotspot Enabled" -Message "The mobile hotspot is now active."
            return $true
        }
        catch {
            LogThis "Hotspot toggle failed: $_" -Color "Red"
            return $false
        }
    }

    # use Devcon.exe to turn Wi-Fi radio off
    function useDevcon {
        
        $devconPath = ".\devcon.exe"
        $deviceId = (Get-PnpDevice -FriendlyName "*Wi-Fi*" | Select-Object -First 1).InstanceId

        Write-Host "Disabling Wi-Fi..."
        & $devconPath disable "$deviceId"
        Start-Sleep -Seconds 3

        Write-Host "Enabling Wi-Fi..."
        & $devconPath enable "$deviceId"

    }    
    
# ==== RUNTIME EXECUTION ====

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

LogThis ""
LogThis "==== Script started @ $timestamp ===="  -Color "Cyan"
LogThis ""

# Optionally force admin mode
if ($forceAsAdmin) {
    LogThis    "Running as Administrator." -verboseMessage $true
    RunAsAdmin
}
if (IsRunningFromTerminal) {
    LogThis "Script is running from Terminal."
    }
else {
    LogThis "Script is running from Task Scheduler. Delaying execution for $delay seconds."
    Start-Sleep -Seconds $delay
    }

# Call the main function
Toggle-Hotspot

LogThis ""
# Keep console open briefly
# Start-Sleep -Seconds 2

Pause
