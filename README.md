## Windows Update Compliance Reporter

PowerShell tool to snapshot Windows Update compliance for a single PC or a fleet.  
Outputs **JSON** and/or **CSV** with pending update counts, last install/detect times, OS build, and a simple compliance flag.

---

## Features

- Local or remote collection (PowerShell Remoting/WinRM)  
- Built-in **Windows Update COM API** (no extra deps) + optional **PSWindowsUpdate** (`-UsePSWindowsUpdate`)  
- **`-HostsCsv`**: point at a CSV/TXT list of machines (no pre-import needed)  
- **`-LogPath`**: per-run log file/dir  
- One-file **HTML report** generator  

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

Run:  

Implicit auth (same domain / rights ok)  
.\Get-UpdateCompliance.ps1 -HostsCsv .\samples\hosts.csv -Csv .\fleet.csv -Json .\fleet.json -LogPath .\logs  

Or explicit creds  
.\Get-UpdateCompliance.ps1 -HostsCsv .\samples\hosts.csv -Credential (Get-Credential) -Csv .\fleet.csv -Json .\fleet.json -LogPath .\logs  

HTML Report:

From your JSON/CSV  
.\New-UpdateComplianceHtml.ps1 -InputPath .\fleet.json -OutHtml .\reports\fleet.html -Open  

Or aggregate many runs  
.\New-UpdateComplianceHtml.ps1 -InputPath .\reports\*.json -OutHtml .\reports\latest.html  

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

-ComputerName <string[]>     # Remote targets; omit for local  
-HostsCsv <path>             # CSV/TXT of targets (header can be ComputerName/Host/Name/FQDN)  
-Credential <pscredential>   # Optional; if omitted, uses current user (Kerberos/NTLM)  
-Json <path>                 # Write JSON  
-Csv <path>                  # Write CSV  
-LogPath <path>              # File or directory for logs  
-MaxDaysSinceInstall <int>   # Default: 14  
-AllowPending                # If set, pending updates won't auto-fail compliance  
-UsePSWindowsUpdate          # Prefer PSWindowsUpdate if available  

---

## Scheduling (optional)

Daily snapshot at 08:00 writing date-stamped files:

$arg = '-NoProfile -ExecutionPolicy Bypass -File "'  + "$PSScriptRoot\Get-UpdateCompliance.ps1" +
       '" -HostsCsv "' + "$PSScriptRoot\samples\hosts.csv" +
       '" -Json "'     + "$PSScriptRoot\reports\$(Get-Date -Format yyyy-MM-dd).json" +
       '" -Csv "'      + "$PSScriptRoot\reports\$(Get-Date -Format yyyy-MM-dd).csv" +
       '" -LogPath "'  + "$PSScriptRoot\logs" + '"'

$action  = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $arg
$trigger = New-ScheduledTaskTrigger -Daily -At 8:00AM
Register-ScheduledTask -TaskName "DAT Update Compliance" -Action $action -Trigger $trigger `
  -Description "Daily Windows Update compliance snapshot"

---

## Troubleshooting

RemoteError / Unknown: verify WinRM (Enable-PSRemoting), firewall rules, DNS, and permissions.  
Workgroup/cross-domain: set TrustedHosts or use -Credential.  
No pending counts: some environments restrict WUA COM; try -UsePSWindowsUpdate (module must be installed on target).  
Times are UTC; convert in your pipeline/UI.  

---



