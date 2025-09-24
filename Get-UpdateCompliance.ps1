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
  [switch]$UsePSWindowsUpdate,
  [string]$LogPath,
  [string]$HostsCsv  
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
  Write-Log INFO "Starting remote compliance on $($Targets.Count) target(s)"

  $sb = {
    param($maxDays,$allowPending,$usePSWU)
    ${function:Get-OSInfo} | Out-Null
    ${function:Get-WULastSuccess} | Out-Null
    ${function:Get-WUState-FromCOM} | Out-Null
    ${function:Get-WUState-FromPSWindowsUpdate} | Out-Null
    ${function:Get-WUComplianceLocal} | Out-Null
    Set-Variable -Name MaxDaysSinceInstall -Value $maxDays -Scope Script
    if ($allowPending) { Set-Variable -Name AllowPending -Value $true -Scope Script }
    if ($usePSWU)      { Set-Variable -Name UsePSWindowsUpdate -Value $true -Scope Script }
    Get-WUComplianceLocal
  }

  $results = @()
  foreach ($t in $Targets) {
    try {
      if ($Cred) {
        $res = Invoke-Command -ComputerName $t -Credential $Cred `
          -ScriptBlock $sb -ArgumentList $MaxDaysSinceInstall,$AllowPending,$UsePSWindowsUpdate -ErrorAction Stop
      } else {
        $res = Invoke-Command -ComputerName $t `
          -ScriptBlock $sb -ArgumentList $MaxDaysSinceInstall,$AllowPending,$UsePSWindowsUpdate -ErrorAction Stop
      }
      $results += $res
      Write-Log INFO  "Remote OK $($t) → Compliance=$($res.Compliance); Pending=$($res.PendingCount); Sec=$($res.PendingSecurity)"
    } catch {
      $msg = $_.Exception.Message
      Write-Log ERROR "RemoteError on $($t): $msg"
      $results += [pscustomobject]@{ ComputerName=$t; Compliance='Unknown'; Reasons="RemoteError: $msg"; CollectedAt=[datetime]::UtcNow }
      }
    }
  $results
}

# --- logging helpers ---

$script:LogFile = $null

function Initialize-Logging {
  param([string]$Path)
  if (-not $Path) { return }
  try {
    if (Test-Path $Path -PathType Leaf) {
      $file = $Path
      $dir  = Split-Path -Parent $file
      if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    } else {
      if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
      $stamp = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH-mm-ssZ')
      $file  = Join-Path $Path "UpdateCompliance_$stamp.log"
    }
    $script:LogFile = $file
    "[$([datetime]::UtcNow.ToString('o'))] INFO  Logging to $file" | Out-File -FilePath $file -Encoding utf8 -Append
  } catch {
    Write-Warning "Failed to init logging: $($_.Exception.Message)"
  }
}

function Write-Log {
  param(
    [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO',
    [Parameter(Mandatory)][string]$Message
  )
  $ts = [datetime]::UtcNow.ToString('o')
  $line = "[$ts] $Level $Message"
  if ($script:LogFile) { $line | Out-File -FilePath $script:LogFile -Encoding utf8 -Append }
  Write-Verbose $line
}

function Get-TargetsFromCsv {
  param([Parameter(Mandatory)][string]$Path)

  if (-not (Test-Path $Path)) { throw "Hosts list not found: $Path" }

  $ext = ([IO.Path]::GetExtension($Path) ?? '').ToLowerInvariant()
  if ($ext -eq '.txt') {
    return (Get-Content -Path $Path | Where-Object { $_ -and $_ -notmatch '^\s*#' } |
            ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique)
  }

  $rows = Import-Csv -Path $Path
  if (-not $rows) { return @() }

  $preferred = @('ComputerName','Hostname','Host','Name','FQDN')
  $props     = ($rows | Get-Member -MemberType NoteProperty | Select-Object -Expand Name) | Sort-Object -Unique
  $col       = ($preferred | Where-Object { $_ -in $props } | Select-Object -First 1)
  if (-not $col) { $col = $props | Select-Object -First 1 }

  $targets = foreach ($r in $rows) {
    $v = $r.$col
    if ($v -and -not [string]::IsNullOrWhiteSpace($v)) { $v.Trim() }
  }
  $targets | Sort-Object -Unique
}


# -------- main --------

$__isDotSourced = $MyInvocation.InvocationName -eq '.'
if (-not $__isDotSourced) {

  Initialize-Logging -Path $LogPath
  Write-Log INFO "Run start. Params: MaxDays=$MaxDaysSinceInstall; AllowPending=$AllowPending; UsePSWU=$UsePSWindowsUpdate"

  # Build target list
  $targets = @()
  if ($HostsCsv) {
    try {
      $fromFile = Get-TargetsFromCsv -Path $HostsCsv
      Write-Log INFO "Loaded $($fromFile.Count) host(s) from $([IO.Path]::GetFileName($HostsCsv))"
      $targets += $fromFile
    } catch {
      Write-Log ERROR "Failed reading HostsCsv: $($_.Exception.Message)"
      throw
    }
  }
  if ($ComputerName) { $targets += $ComputerName }
  $targets = $targets | Sort-Object -Unique

  $all = @()
  if ($targets.Count -gt 0) {
    if (-not $Credential) { Write-Log INFO "No -Credential provided; attempting implicit auth." }
    Write-Log INFO "Targets: $($targets -join ', ')"
    $all = Invoke-RemoteCompliance -Targets $targets -Cred $Credential
  } else {
    $all = @(Get-WUComplianceLocal)
  }

  foreach ($r in $all) {
    Write-Log INFO ("Summary {0} → {1}; Pending={2}; Sec={3}; LastInstall={4}" -f $r.ComputerName,$r.Compliance,$r.PendingCount,$r.PendingSecurity,$r.LastInstallSuccess)
  }

  $all | Select-Object ComputerName,Compliance,Reasons,PendingCount,PendingSecurity,LastInstallSuccess,DaysSinceInstall,Build |
    Sort-Object ComputerName | Format-Table -AutoSize

  if ($Json) {
    $all | ConvertTo-Json -Depth 6 | Out-File -Encoding utf8 $Json
    Write-Log INFO "Wrote JSON -> $Json"
  }
  if ($Csv) {
    $all | Select-Object ComputerName,Compliance,Reasons,PendingCount,PendingSecurity,LastDetectSuccess,LastInstallSuccess,DaysSinceInstall,ProductName,Edition,DisplayVersion,Build,CollectedAt |
      Export-Csv -NoTypeInformation -Encoding UTF8 $Csv
    Write-Log INFO "Wrote CSV -> $Csv"
  }

  Write-Log INFO "Run complete. Hosts=$($all.Count)"
}


