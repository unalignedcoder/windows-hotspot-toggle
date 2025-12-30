<#
.SYNOPSIS
    Hotspot Toggle with Persistent WiFi Adapter Selection
.DESCRIPTION
    A modern PowerShell script to toggle the WiFi hotspot on Windows 10/11
.NOTES
    rehauled main function
    minor fixes
#>

# FORCE POWERSHELL 5.1 HANDOFF (Must come first to avoid Type errors in PS7)
if ($PSVersionTable.PSVersion.Major -gt 5) {
    # Check if we can even log this yet
    Write-Host "Detected PowerShell 7+. Relaunching in PowerShell 5.1 for WinRT compatibility..." -ForegroundColor Cyan
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    Exit
}

# LOAD ASSEMBLIES & NAMESPACES
try {
    Add-Type -AssemblyName "System.Runtime.WindowsRuntime" -ErrorAction Stop
    $WinRTAssemblies = @(
        "Windows.Networking.Connectivity.NetworkInformation, Windows.Networking.Connectivity, ContentType=WindowsRuntime",
        "Windows.Networking.NetworkOperators.NetworkOperatorTetheringManager, Windows.Networking.NetworkOperators, ContentType=WindowsRuntime",
        "Windows.Devices.Radios.Radio, Windows.System.Devices, ContentType=WindowsRuntime",
        "Windows.Devices.Radios.RadioAccessStatus, Windows.System.Devices, ContentType=WindowsRuntime",
        "Windows.Devices.Radios.RadioState, Windows.System.Devices, ContentType=WindowsRuntime"
    )
    foreach ($asm in $WinRTAssemblies) {
        [void][System.Runtime.InteropServices.Marshal]::PrelinkAll([type]::GetType($asm))
    }
} catch {
    Write-Warning "WinRT Type loading encountered an issue: $_"
}

# ==== Configuration ====
    $scriptVersion = "1.0.13"
    $configFile    = "$PSScriptRoot\adapter.config"
    $logFile       = $true
    $logReverse    = $false
    $logFilePath   = "$PSScriptRoot\script.log"
    $restartWiFi   = $false 
    $forceAsAdmin  = $true
    $delay         = 30

# ==== WinRT Helper Logic ====
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

    # Helper to await WinRT async operations
    function Await-WinRT {
        param ($winRtOperation, [Type]$resultType)
        try {
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

    # BurnToast Notification Functionality
    function Send-Notification {
        param ([string]$Title, [string]$Message)
        if (-not (IsRunningFromTerminal)) {
            if (Get-Module -ListAvailable -Name BurntToast) {
                New-BurntToastNotification -Text $Title, $Message
            }
        }
    }

    # Admin Check & Relaunch
    function IsAdmin {
        $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }

    # Relaunch script as admin if not already
    function RunAsAdmin {
        if ($env:SYSTEMROOT -and $env:USERNAME -eq "SYSTEM") { return }
        if (-Not (IsAdmin)) {
            Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
            Exit 0
        }
    }

    # Logging Functionality
    function LogThis {
        param ([string]$Message, [string]$Color = "White")
        Write-Host $Message -ForegroundColor $Color
        if ($logFile) {
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMsg = "[$timestamp] $Message"
            if ($logReverse) {
                $existing = if (Test-Path $logFilePath) { Get-Content $logFilePath -Raw } else { "" }
                "$logMsg`n$existing" | Set-Content $logFilePath -Encoding UTF8
            } else {
                $logMsg | Out-File -Append $logFilePath -Encoding UTF8
            }
        }
    }

    # Check if running from a terminal
    function IsRunningFromTerminal {
        $proc = Get-CimInstance Win32_Process -Filter "ProcessId = $pid"
        $parent = Get-CimInstance Win32_Process -Filter "ProcessId = $($proc.ParentProcessId)"
        return -not ($parent.Name -eq "svchost.exe" -or $parent.CommandLine -like "*Schedule*")
    }
    
    # Get or select WiFi adapter
    function Get-WiFiAdapter {
        if (Test-Path $configFile) {
            try {
                $saved = Get-Content $configFile | ConvertFrom-Json
                $adapter = Get-NetAdapter | Where-Object { $_.Name -eq $saved.Name -and $_.InterfaceDescription -eq $saved.Description }
                if ($adapter) { return $adapter }
            } catch { LogThis "Config error: $_" -Color "Yellow" }
        }
        $adapter = Get-NetAdapter | Where-Object { $_.PhysicalMediaType -eq 'Native 802.11' } | Select-Object -First 1
        if ($adapter) { return $adapter }
        
        $selected = Get-NetAdapter | Out-GridView -Title "Select WiFi Adapter" -PassThru
        if ($selected) {
            [PSCustomObject]@{Name=$selected.Name; Description=$selected.InterfaceDescription} | ConvertTo-Json | Out-File $configFile
            return $selected
        }
        throw "No WiFi adapter found"
    }

    # Set WiFi Radio State
    function Set-WifiRadioState {
        param ([string]$TargetState)
        try {
            Await-WinRT ([Windows.Devices.Radios.Radio]::RequestAccessAsync()) ([Windows.Devices.Radios.RadioAccessStatus]) | Out-Null
            $radios = Await-WinRT ([Windows.Devices.Radios.Radio]::GetRadiosAsync()) ([System.Collections.Generic.IReadOnlyList[Windows.Devices.Radios.Radio]])
            $wifiRadio = $radios | Where-Object { $_.Kind -eq 'WiFi' }
            if ($wifiRadio) {
                $state = if ($TargetState -eq "On") { [Windows.Devices.Radios.RadioState]::On } else { [Windows.Devices.Radios.RadioState]::Off }
                LogThis "Radio: $TargetState" -Color "Yellow"
                Await-WinRT ($wifiRadio.SetStateAsync($state)) ([Windows.Devices.Radios.RadioAccessStatus]) | Out-Null
                return $true
            }
        } catch { LogThis "Radio error: $_" -Color "Red" }
        return $false
    }

    # Main Hotspot Toggle Logic
    function Switch-Hotspot {
        param($adapter)
        try {
            $connectionProfile = $null
            $retryCount = 0
            $maxRetries = 12 # Up to 60 seconds

            while ($null -eq $connectionProfile -and $retryCount -lt $maxRetries) {
                $connectionProfile = [Windows.Networking.Connectivity.NetworkInformation]::GetInternetConnectionProfile()
                if ($null -eq $connectionProfile) {
                    LogThis "Waiting for Internet profile (Attempt $($retryCount + 1)/$maxRetries)..." -Color "Yellow"
                    Start-Sleep -Seconds 5
                    $retryCount++
                }
            }

            if ($null -eq $connectionProfile) {
                LogThis "Error: No Internet profile found. Aborting." -Color "Red"
                return $false
            }

            $tetheringManager = [Windows.Networking.NetworkOperators.NetworkOperatorTetheringManager]::CreateFromConnectionProfile($connectionProfile)

            if ($tetheringManager.TetheringOperationalState -eq 'On') {
                LogThis "Hotspot ON -> Disabling..." -Color "Yellow"
                $null = Await-WinRT ($tetheringManager.StopTetheringAsync()) ([Windows.Networking.NetworkOperators.NetworkOperatorTetheringOperationResult])
                Send-Notification -Title "Hotspot" -Message "Disabled"
                return $true
            }

            LogThis "Resetting Radio for Win10 Routing Fix..." -Color "Cyan"
            Set-WifiRadioState -TargetState "Off"
            Start-Sleep -Seconds 2
            Set-WifiRadioState -TargetState "On"
            Start-Sleep -Seconds 6 # Added more time for Win10 DHCP stability

            if ($restartWiFi) {
                LogThis "Restarting Adapter: $($adapter.Name)" -Color "Yellow"
                Restart-NetAdapter -Name $adapter.Name -Confirm:$false
                Start-Sleep -Seconds 3
            }

            LogThis "Enabling Hotspot..." -Color "Yellow"
            $result = Await-WinRT ($tetheringManager.StartTetheringAsync()) ([Windows.Networking.NetworkOperators.NetworkOperatorTetheringOperationResult])
            if ($result.Status -eq "Success") {
                LogThis "Success! Hotspot is active." -Color "Green"
                Send-Notification -Title "Hotspot" -Message "Enabled"
            } else {
                LogThis "Failed. Status: $($result.Status)" -Color "Red"
            }
        } catch { LogThis "Critical error: $_" -Color "Red" }
    }

# ==== RUNTIME EXECUTION ====

LogThis "==== Script Execution Initiated ====" -Color "Cyan"
LogThis "Script Version: $scriptVersion"

if ($forceAsAdmin) { RunAsAdmin }

$targetAdapter = Get-WiFiAdapter

if (-not (IsRunningFromTerminal)) {
    LogThis "Background mode: Delaying $delay sec for system stability..."
    Start-Sleep -Seconds $delay
}

Switch-Hotspot -adapter $targetAdapter
LogThis "==== Done ===="
