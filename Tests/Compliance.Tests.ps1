$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Resolve-Path "$here\.."

# Dot-source without running main (guard was added)
. "$root\Get-UpdateCompliance.ps1"

Describe 'Get-WUComplianceLocal' {
  BeforeAll {
    $here = Split-Path -Parent $PSCommandPath
    $root = Split-Path -Parent $here
    . "$root\Get-UpdateCompliance.ps1"   # load functions at run time
    $script:UsePSWindowsUpdate = $false  # keep COM path in tests
  }

  It 'marks NonCompliant when last install > MaxDays and pending > 0' {
    $script:MaxDaysSinceInstall = 14
    $script:AllowPending = $false

    Mock Get-OSInfo { [pscustomobject]@{
      ComputerName='TEST01'; ProductName='Windows 11'; Edition='Pro'
      DisplayVersion='24H2'; Build='26100.1234'
    }}

    Mock Get-WULastSuccess { [pscustomobject]@{
      LastInstallSuccess = (Get-Date).AddDays(-20)
      LastDetectSuccess  = (Get-Date).AddDays(-1)
    }}

    Mock Get-WUState-FromCOM { [pscustomobject]@{
      PendingCount=3; PendingSecurity=2; PendingUpdates=@(); LastInstallFromHist=$null
    }}

    $res = Get-WUComplianceLocal
    $res.Compliance | Should -Be 'NonCompliant'
    $res.Reasons    | Should -Match 'LastInstall>14'
    $res.Reasons    | Should -Match 'PendingUpdates=3'
  }

  It 'marks Compliant when recent install and no pending' {
    $script:MaxDaysSinceInstall = 14
    $script:AllowPending = $false

    Mock Get-OSInfo { [pscustomobject]@{
      ComputerName='TEST02'; ProductName='Windows 10'; Edition='Pro'
      DisplayVersion='22H2'; Build='19045.4529'
    }}

    Mock Get-WULastSuccess { [pscustomobject]@{
      LastInstallSuccess = (Get-Date).AddDays(-2)
      LastDetectSuccess  = (Get-Date)
    }}

    Mock Get-WUState-FromCOM { [pscustomobject]@{
      PendingCount=0; PendingSecurity=0; PendingUpdates=@(); LastInstallFromHist=$null
    }}

    $res = Get-WUComplianceLocal
    $res.Compliance | Should -Be 'Compliant'
    $res.Reasons    | Should -Be ''
  }
}