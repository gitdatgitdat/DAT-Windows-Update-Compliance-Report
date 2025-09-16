## Windows Update Compliance Reporter

Lightweight PowerShell tool to snapshot Windows Update compliance for a single PC or a fleet.  
Outputs **JSON** and/or **CSV** with pending update counts, last install/detect times, OS build, and a simple compliance flag.

---

## Features
- Local or remote collection (PowerShell Remoting/WinRM)
- Uses built-in **Windows Update COM API** (no extra deps)
- Optional **PSWindowsUpdate** path (`-UsePSWindowsUpdate`)
- JSON & CSV exports for dashboards or SIEM
- Clear compliance reasons (e.g., `LastInstall>14 days; PendingUpdates=3`)

---

## Quick Start

powershell  
Set-ExecutionPolicy -Scope Process RemoteSigned  

Local run (console preview):    
.\Get-UpdateCompliance.ps1  

Export JSON & CSV:    
.\Get-UpdateCompliance.ps1 -Json .\report.json -Csv .\report.csv  

Tighter compliance window:    
.\Get-UpdateCompliance.ps1 -MaxDaysSinceInstall 7  

Multi-host:    
Create samples/hosts.csv:  

ComputerName  
PC1  
PC2  
PC3.domain.local  

Run with creds:  

$cred  = Get-Credential  
$hosts = (Import-Csv .\samples\hosts.csv).ComputerName  
.\Get-UpdateCompliance.ps1 -ComputerName $hosts -Credential $cred -Csv .\fleet.csv -Json .\fleet.json  

Requirements: WinRM enabled & reachable, account allowed to connect.  

---

## Compliance Logic (default)

Rule | Default | Notes  
Last successful install older than -MaxDaysSinceInstall |	14 days |	Non-compliant  
Any pending updates	| N/A	| Non-compliant unless -AllowPending  

Adjust with -MaxDaysSinceInstall and/or -AllowPending.

--- 

## Output Fields (per machine)

Key columns/props:

Compliance: Compliant | NonCompliant | Unknown  
Reasons: semicolon list, e.g. LastInstall>14 days; PendingUpdates=3  
PendingCount, PendingSecurity  
LastDetectSuccess, LastInstallSuccess, DaysSinceInstall  
ProductName, Edition, DisplayVersion, Build  
CollectedAt (UTC)  

Example JSON: samples/example-output.json

---

## Parameters

-ComputerName <string[]>   # Remote targets; omit for local  
-Credential <pscredential> # Required with -ComputerName  
-Json <path>               # Write JSON  
-Csv <path>                # Write CSV  
-MaxDaysSinceInstall <int> # Default: 14  
-AllowPending              # If set, pending updates won't auto-fail compliance  
-UsePSWindowsUpdate        # Prefer PSWindowsUpdate if available  

---

## Scheduling (optional)

Daily snapshot at 08:00 writing date-stamped files:

$action  = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$PSScriptRoot\Get-UpdateCompliance.ps1`" -Json `"$PSScriptRoot\reports\$(Get-Date -Format yyyy-MM-dd).json`" -Csv `"$PSScriptRoot\reports\$(Get-Date -Format yyyy-MM-dd).csv`""
$trigger = New-ScheduledTaskTrigger -Daily -At 8:00AM
Register-ScheduledTask -TaskName "DAT Update Compliance" -Action $action -Trigger $trigger -Description "Daily Windows Update compliance snapshot"

---

## Troubleshooting

RemoteError / Unknown: verify WinRM (Enable-PSRemoting), firewall, DNS, and creds.  
No pending counts: some environments restrict WUA COM; try -UsePSWindowsUpdate (module must be installed on target).  
Dates: times are UTC; convert as needed in your pipeline/UI.  

---
