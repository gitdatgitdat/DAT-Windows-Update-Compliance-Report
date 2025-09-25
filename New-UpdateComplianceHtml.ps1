[CmdletBinding()]
param(
  # One or more input files: JSON or CSV (wildcards OK)
  [Parameter(Mandatory)][string[]]$InputPath,
  # Output HTML file
  [string]$OutHtml = ".\reports\UpdateCompliance.html",
  # Open the report when done
  [switch]$Open
)

# --- helpers ---
function Read-Records {
  param([string]$Path)
  $files = @()
  foreach ($p in $Path) { $files += Get-ChildItem -Path $p -File -ErrorAction Stop }
  if (-not $files) { throw "No files matched: $Path" }

  $all = @()
  foreach ($f in $files) {
    switch ($f.Extension.ToLower()) {
      '.json' { $all += @((Get-Content -Raw $f.FullName | ConvertFrom-Json)) }
      '.csv'  { $all += @(Import-Csv $f.FullName) }
      default { Write-Warning "Skipping unsupported file: $($f.Name)" }
    }
  }
  # Normalize and dedupe by latest CollectedAt per ComputerName
  $norm = foreach ($r in $all) {
    [pscustomobject]@{
      ComputerName       = $r.ComputerName
      Compliance         = $r.Compliance
      Reasons            = $r.Reasons
      PendingCount       = [int]$r.PendingCount
      PendingSecurity    = [int]$r.PendingSecurity
      LastInstallSuccess = if ($r.LastInstallSuccess) { [datetime]$r.LastInstallSuccess } else { $null }
      DaysSinceInstall   = if ($r.DaysSinceInstall -ne $null -and $r.DaysSinceInstall -ne '') { [int]$r.DaysSinceInstall } else { $null }
      Build              = $r.Build
      ProductName        = $r.ProductName
      CollectedAt        = if ($r.CollectedAt) { [datetime]$r.CollectedAt } else { Get-Date }
    }
  }

  $latest = $norm |
    Where-Object { $_.ComputerName } |
    Group-Object ComputerName |
    ForEach-Object { $_.Group | Sort-Object CollectedAt -Descending | Select-Object -First 1 }

  ,$latest
}

function HtmlEncode([string]$s) {
  if ($null -eq $s) { return '' }
  $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
}

function Dt([datetime]$d) {
  if ($d) { $d.ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { '' }
}

# --- data ---
$data = Read-Records -Path $InputPath
$total = $data.Count
$counts = @{
  Compliant     = ($data | Where-Object { $_.Compliance -eq 'Compliant' }).Count
  NonCompliant  = ($data | Where-Object { $_.Compliance -eq 'NonCompliant' }).Count
  Unknown       = ($data | Where-Object { $_.Compliance -eq 'Unknown' }).Count
}
$generated = (Get-Date).ToString('yyyy-MM-dd HH:mm')

# --- rows ---
$rows = foreach ($r in ($data | Sort-Object ComputerName)) {
  $cls = switch ($r.Compliance) {
    'Compliant'     { 'ok' }
    'NonCompliant'  { 'bad' }
    default         { 'unk' }
  }
  @"
<tr class="$cls">
  <td class="name">$(HtmlEncode $r.ComputerName)</td>
  <td class="status"><span class="dot"></span>$(HtmlEncode $r.Compliance)</td>
  <td class="num">$($r.PendingCount)</td>
  <td class="num">$($r.PendingSecurity)</td>
  <td>$(Dt $r.LastInstallSuccess)</td>
  <td class="num">$($r.DaysSinceInstall)</td>
  <td>$(HtmlEncode $r.Build)</td>
  <td class="reasons">$(HtmlEncode $r.Reasons)</td>
</tr>
"@
}

# --- html ---
$html = @"
<!doctype html>
<html lang="en">
<meta charset="utf-8"/>
<title>Windows Update Compliance</title>
<style>
  :root { --ok:#22c55e; --bad:#ef4444; --unk:#9ca3af; --ink:#111827; --muted:#6b7280; --bg:#ffffff; --row:#f9fafb; }
  body { font-family: ui-sans-serif,system-ui,Segoe UI,Roboto,Arial; margin:24px; color:var(--ink); background:var(--bg);}
  h1 { margin:0 0 4px 0; font-size:24px; }
  .sub { color:var(--muted); margin-bottom:16px; }
  .cards { display:flex; gap:12px; margin:14px 0 18px; }
  .card { padding:10px 12px; border:1px solid #e5e7eb; border-radius:10px; box-shadow:0 1px 2px rgba(0,0,0,.04); }
  .card strong { font-size:18px; margin-right:6px }
  .controls { display:flex; gap:12px; margin:10px 0 14px; align-items:center; }
  select,input[type=search]{ padding:6px 8px; border:1px solid #e5e7eb; border-radius:8px; }
  table { width:100%; border-collapse:collapse; }
  th,td { padding:10px 8px; border-bottom:1px solid #e5e7eb; }
  tr:nth-child(even){ background:var(--row); }
  th { text-align:left; font-weight:600; }
  .num { text-align:right; }
  .status { white-space:nowrap; font-weight:600; }
  .name { font-weight:600; }
  .dot { display:inline-block; width:10px; height:10px; border-radius:50%; margin-right:6px; vertical-align:middle; }
  tr.ok  .dot { background:var(--ok); }
  tr.bad .dot { background:var(--bad); }
  tr.unk .dot { background:var(--unk); }
  .reasons { max-width:520px; overflow-wrap:anywhere; }
  .muted { color:var(--muted); }
  .pill { padding:2px 8px; border-radius:9999px; font-size:12px; border:1px solid #e5e7eb; }
</style>

<h1>Windows Update Compliance</h1>
<div class="sub">Generated $generated</div>

<div class="cards">
  <div class="card"><strong>$total</strong><span class="muted">Total</span></div>
  <div class="card"><strong style="color:var(--ok)">$($counts.Compliant)</strong><span class="muted">Compliant</span></div>
  <div class="card"><strong style="color:var(--bad)">$($counts.NonCompliant)</strong><span class="muted">Non-compliant</span></div>
  <div class="card"><strong style="color:var(--unk)">$($counts.Unknown)</strong><span class="muted">Unknown</span></div>
</div>

<div class="controls">
  <label>Filter:</label>
  <select id="flt" onchange="applyFilter()">
    <option value="all" selected>All</option>
    <option value="ok">Compliant</option>
    <option value="bad">Non-compliant</option>
    <option value="unk">Unknown</option>
  </select>
  <input id="q" type="search" placeholder="Search computer or reason…" oninput="applyFilter()"/>
  <span class="pill muted">Click headers to sort</span>
</div>

<table id="tbl">
  <thead>
    <tr>
      <th onclick="sortBy(0)">Computer</th>
      <th onclick="sortBy(1)">Compliance</th>
      <th class="num" onclick="sortBy(2)">Pending</th>
      <th class="num" onclick="sortBy(3)">Security</th>
      <th onclick="sortBy(4)">Last Install</th>
      <th class="num" onclick="sortBy(5)">Days</th>
      <th onclick="sortBy(6)">Build</th>
      <th onclick="sortBy(7)">Reasons</th>
    </tr>
  </thead>
  <tbody>
$($rows -join "")
  </tbody>
</table>

<script>
let asc=true;
function sortBy(col){
  const tbody=document.querySelector('#tbl tbody');
  const rows=[...tbody.querySelectorAll('tr')];
  rows.sort((a,b)=>{
    const A=a.children[col].innerText.trim().toLowerCase();
    const B=b.children[col].innerText.trim().toLowerCase();
    const nA=parseFloat(A); const nB=parseFloat(B);
    const isNum=!isNaN(nA) && !isNaN(nB);
    if(isNum){ return asc ? nA-nB : nB-nA; }
    return asc ? A.localeCompare(B) : B.localeCompare(A);
  });
  asc=!asc; rows.forEach(r=>tbody.appendChild(r));
}
function applyFilter(){
  const val=document.getElementById('flt').value;
  const q=document.getElementById('q').value.toLowerCase();
  document.querySelectorAll('#tbl tbody tr').forEach(tr=>{
    const matchesClass = (val==='all') || tr.classList.contains(val);
    const text = tr.innerText.toLowerCase();
    const matchesSearch = !q || text.includes(q);
    tr.style.display = (matchesClass && matchesSearch) ? '' : 'none';
  });
}
</script>
</html>
"@

# write file
$dir = Split-Path -Parent $OutHtml
if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
$html | Out-File -FilePath $OutHtml -Encoding UTF8
Write-Host "[OK] Wrote HTML -> $OutHtml"
if ($Open) { Start-Process $OutHtml | Out-Null }
