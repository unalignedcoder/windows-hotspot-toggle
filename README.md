# Windows Hotspot Toggle  

PowerShell script to toggle the WiFi hotspot on Windows 10/11 without using deprecated `netsh` commands.  

## Features  

This script is meant to automatically enable or disable the Windows Hotspot.
In order for it to properly work and be firewall-firendly, some workarounds have been included:
the wifi radio is turned off and on, optionally the adpater itself is reset before operation.

In our tests this should work at startup as well, so as to have the hotspot always running. 

## Usage

### Check configuration segment in the script

<img width="692" height="272" alt="image" src="https://github.com/user-attachments/assets/4514fff1-2f2e-497a-a62b-4882d2480bc3" />

### Run script in terminal
`.\toggle-hotspot.ps1`

### Run as a shortcut, or in Task Scheduler
`%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe -file toggle-hotspot.ps1`

## Requirements
- Windows 10/11  
- Administrator privileges  
- PowerShell 5.1+  
