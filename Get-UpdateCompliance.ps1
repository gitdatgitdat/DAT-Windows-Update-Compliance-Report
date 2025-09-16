<#
.SYNOPSIS
  Windows Update Compliance Reporter (local + optional multi-host)

.DESCRIPTION
  Collects Windows Update state:
    - Pending updates (count, security count)
    - Installed update history (last install time)
    - Last successful detection (scan) time
    - OS version/build
  Flags compliance if:
    - Last successful install older than -MaxDaysSinceInstall
    - Any pending updates exist (unless -AllowPending is set)

.PARAMETER ComputerName
  One or more remote computers. If omitted, runs on the local machine.

.PARAMETER Credential
  Credential for remote execution (when ComputerName is used).

.PARAMETER Json
  Path to write JSON report.

.PARAMETER Csv
  Path to write CSV report (one row per machine).

.PARAMETER MaxDaysSinceInstall
  Threshold in days; if last install is older, mark NonCompliant. Default: 14.

.PARAMETER AllowPending
  If set, pending updates do not automatically mark NonCompliant.

.PARAMETER UsePSWindowsUpdate
  Prefer the community PSWindowsUpdate module if available (optional).

.EXAMPLE
  # Local JSON + CSV
  .\Get-UpdateCompliance.ps1 -Json report.json -Csv report.csv

.EXAMPLE
  # Remote against a list
  $cred = Get-Credential
  .\Get-UpdateCompliance.ps1 -ComputerName PC1,PC2 -Credential $cred -Csv fleet.csv
#>

[CmdletBinding()]
param(
  [string[]]$ComputerName,
  [pscredential]$Credential,
  [string]$Json,
  [string]$Csv,
  [int]$MaxDaysSinceInstall = 14,
  [switch]$AllowPending,
  [switch]$UsePSWindowsUpdate
)

function Get-OSInfo {
  $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
  $build = "{0}.{1}" -f $cv.CurrentBuildNumber, ($cv.UBR ?? 0)
  [pscustomobject]@{
    ComputerName   = $env:COMPUTERNAME
    Edition        = $cv.EditionID
    DisplayVersion = $cv.DisplayVersion
    ReleaseId      = $cv.ReleaseId
    Build          = $build
    ProductName    = $cv.ProductName
  }
}

function Get-WULastSuccess {
  $paths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Install',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Detect',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Detection' # rare older key
  )
  $out = [ordered]@{ LastInstallSuccess = $null; LastDetectSuccess = $null }
  foreach ($p in $paths) {
    try {
      if (Test-Path $p) {
        $v = Get-ItemProperty $p
        $ts = $v.LastSuccessTime
        if ($ts) {
          $dt = [datetime]::Parse($ts)
          if ($p -like '*\Install')   { $out['LastInstallSuccess'] = $dt }
          if ($p -like '*\Detect*')   { $out['LastDetectSuccess']  = $dt }
        }
      }
    } catch { }
  }
  [pscustomobject]$out
}

function Get-WUState-FromCOM {
  # Windows Update Agent COM (no extra modules)
  try {
    $session  = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()

    $pending = $searcher.Search("IsInstalled=0 and IsHidden=0 and Type='Software'")
    $pendingUpdates      = @()
    $pendingSecurityOnly = 0
    if ($pending.ResultCode -ne 2) { } # 2 = Succeeded
    foreach ($u in @($pending.Updates)) {
      $isSecurity = $false
      foreach ($c in @($u.Categories)) {
        if ($c.Name -match 'Security') { $isSecurity = $true; break }
      }
      if ($isSecurity) { $pendingSecurityOnly++ }
      $pendingUpdates += [pscustomobject]@{
        Title     = $u.Title
        KB        = ($u.KBArticleIDs -join ',')
        IsSecurity= $isSecurity
        Severity  = $u.MsrcSeverity
        Deadline  = $u.Deadline
      }
    }

    # Optional: last install from history
    $lastInstall = $null
    try {
      $count = $searcher.GetTotalHistoryCount()
      if ($count -gt 0) {
        $hist = $searcher.QueryHistory(0, [Math]::Min($count, 50))
        # Operation: 1=Installation, 2=Uninstallation
        $inst = @($hist | Where-Object { $_.Operation -eq 1 } | Sort-Object -Property Date -Descending)[0]
        if ($inst) { $lastInstall = $inst.Date }
      }
    } catch { }

    [pscustomobject]@{
      PendingCount        = @($pendingUpdates).Count
      PendingSecurity     = $pendingSecurityOnly
      PendingUpdates      = $pendingUpdates
      LastInstallFromHist = $lastInstall
    }
  } catch {
    $null
  }
}

function Get-WUState-FromPSWindowsUpdate {
  # Optional path via PSWindowsUpdate module (if installed)
  try {
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) { return $null }
    Import-Module PSWindowsUpdate -ErrorAction Stop | Out-Null
    $avail = Get-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot -NotCategory 'Drivers' -ErrorAction SilentlyContinue
    $pending = @($avail | Where-Object {$_.IsDownloaded -or $_.KB})
    $secOnly = @($pending | Where-Object { $_.Categories -match 'Security' })
    [pscustomobject]@{
      PendingCount    = $pending.Count
      PendingSecurity = $secOnly.Count
      PendingUpdates  = @($pending | Select-Object Title, KB, Categories)
    }
  } catch {
    $null
  }
}

function Get-WUComplianceLocal {
  $os   = Get-OSInfo
  $tim  = Get-WULastSuccess
  $wu   = $null

  if ($UsePSWindowsUpdate) {
    $wu = Get-WUState-FromPSWindowsUpdate
  }
  if (-not $wu) {
    $wu = Get-WUState-FromCOM
  }

  $lastInstall = $tim.LastInstallSuccess
  if (-not $lastInstall -and $wu.LastInstallFromHist) {
    $lastInstall = $wu.LastInstallFromHist
  }

  $daysSinceInstall = $null
  if ($lastInstall) {
    $daysSinceInstall = [int]([datetime]::UtcNow - $lastInstall.ToUniversalTime()).TotalDays
  }

  $pendingCount    = $wu.PendingCount
  $pendingSecurity = $wu.PendingSecurity

  $nonCompliantReasons = @()
  if ($daysSinceInstall -ne $null -and $daysSinceInstall -gt $MaxDaysSinceInstall) {
    $nonCompliantReasons += "LastInstall>$MaxDaysSinceInstall days"
  }
  if (-not $AllowPending -and ($pendingCount -gt 0)) {
    $nonCompliantReasons += "PendingUpdates=$pendingCount"
  }

  [pscustomobject]@{
    ComputerName        = $os.ComputerName
    ProductName         = $os.ProductName
    Edition             = $os.Edition
    DisplayVersion      = $os.DisplayVersion
    Build               = $os.Build
    LastDetectSuccess   = $tim.LastDetectSuccess
    LastInstallSuccess  = $lastInstall
    DaysSinceInstall    = $daysSinceInstall
    PendingCount        = $pendingCount
    PendingSecurity     = $pendingSecurity
    Compliance          = if ($nonCompliantReasons.Count) { "NonCompliant" } else { "Compliant" }
    Reasons             = ($nonCompliantReasons -join "; ")
    PendingUpdates      = $wu.PendingUpdates  # detailed list
    CollectedAt         = [datetime]::UtcNow
  }
}

function Invoke-RemoteCompliance {
  param([string[]]$Targets, [pscredential]$Cred)

  $sb = {
    param($maxDays,$allowPending,$usePSWU)
    # Re-import the functions needed on the remote side:
    ${function:Get-OSInfo} | Out-Null
    ${function:Get-WULastSuccess} | Out-Null
    ${function:Get-WUState-FromCOM} | Out-Null
    ${function:Get-WUState-FromPSWindowsUpdate} | Out-Null
    ${function:Get-WUComplianceLocal} | Out-Null

    # Bind the script-level params
    Set-Variable -Name MaxDaysSinceInstall -Value $maxDays -Scope Script
    if ($allowPending) { Set-Variable -Name AllowPending -Value $true -Scope Script }
    if ($usePSWU)      { Set-Variable -Name UsePSWindowsUpdate -Value $true -Scope Script }

    Get-WUComplianceLocal
  }

  $results = @()
  foreach ($t in $Targets) {
    try {
      $res = Invoke-Command -ComputerName $t -Credential $Cred -ScriptBlock $sb -ArgumentList $MaxDaysSinceInstall,$AllowPending,$UsePSWindowsUpdate -ErrorAction Stop
      $results += $res
    } catch {
      $results += [pscustomobject]@{
        ComputerName       = $t
        Compliance         = "Unknown"
        Reasons            = "RemoteError: $($_.Exception.Message)"
        CollectedAt        = [datetime]::UtcNow
      }
    }
  }
  $results
}

# -------- main --------

$all = @()
if ($ComputerName) {
  if (-not $Credential) {
    Write-Error "Please supply -Credential when using -ComputerName."
    exit 1
  }
  $all = Invoke-RemoteCompliance -Targets $ComputerName -Cred $Credential
} else {
  $all = @(Get-WUComplianceLocal)
}

# Console preview
$all | Select-Object ComputerName,Compliance,Reasons,PendingCount,PendingSecurity,LastInstallSuccess,DaysSinceInstall,Build |
  Sort-Object ComputerName |
  Format-Table -AutoSize

# Outputs
if ($Json) {
  $all | ConvertTo-Json -Depth 6 | Out-File -Encoding utf8 $Json
  Write-Host "[OK] Wrote JSON -> $Json"
}
if ($Csv) {
  $all |
    Select-Object ComputerName,Compliance,Reasons,PendingCount,PendingSecurity,LastDetectSuccess,LastInstallSuccess,DaysSinceInstall,ProductName,Edition,DisplayVersion,Build,CollectedAt |
    Export-Csv -NoTypeInformation -Encoding UTF8 $Csv
  Write-Host "[OK] Wrote CSV  -> $Csv"
}
