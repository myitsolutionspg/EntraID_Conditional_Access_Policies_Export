# Entra Conditional Access – Dynamic HTML Report
# Requires Microsoft.Graph and at least Policy.Read.All.
# Recommended scopes: Policy.Read.All, Directory.Read.All, Organization.Read.All

Connect-MgGraph -Scopes "Policy.Read.All","Directory.Read.All","Organization.Read.All"

# ---------- Helper functions ----------
function Encode-Html {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    $t = $Text
    $t = $t -replace '&', '&amp;'
    $t = $t -replace '<', '&lt;'
    $t = $t -replace '>', '&gt;'
    return $t
}

function Format-List {
    param($Values)
    if (-not $Values) { return "" }
    if ($Values -is [string]) { return (Encode-Html $Values) }
    return ($Values | ForEach-Object { Encode-Html $_ }) -join "<br/>"
}

function Format-ObjShort {
    param($Obj)
    if ($null -eq $Obj) { return "" }
    if ($Obj -is [string] -or $Obj -is [ValueType]) { return (Encode-Html $Obj.ToString()) }
    $json = $Obj | ConvertTo-Json -Depth 5 -Compress
    return (Encode-Html $json)
}

function Format-SessionObject {
    param($Obj)

    if ($null -eq $Obj) {
        return "<span class='muted'>Not configured</span>"
    }

    $lines = @()

    foreach ($prop in $Obj.PSObject.Properties) {
        if ($prop.Name -eq "AdditionalProperties") { continue }

        $value = $prop.Value

        if ($null -eq $value -or ($value -is [string] -and $value -eq "")) { continue }

        if ($value -is [bool]) {
            $value = if ($value) { "Yes" } else { "No" }
        }

        $lines += ("{0}: {1}" -f $prop.Name, $value)
    }

    if (-not $lines) {
        return "<span class='muted'>Not configured</span>"
    }

    return ($lines | ForEach-Object { Encode-Html $_ }) -join "<br/>"
}

function Format-Platforms {
    param($Platforms)

    if ($null -eq $Platforms) {
        return "<span class='muted'>Not configured</span>"
    }

    $lines = @()

    $include = $Platforms.IncludePlatforms
    $exclude = $Platforms.ExcludePlatforms

    if ($include) {
        if (($include -is [System.Collections.IEnumerable]) -and ($include -contains "all")) {
            $lines += "Include: All platforms"
        } else {
            $lines += "Include: " + ($include -join ", ")
        }
    }

    if ($exclude) {
        $lines += "Exclude: " + ($exclude -join ", ")
    }

    if (-not $lines) {
        return "<span class='muted'>Not configured</span>"
    }

    return ($lines | ForEach-Object { Encode-Html $_ }) -join "<br/>"
}

function Format-DeviceFilter {
    param($DevicesCondition)

    if ($null -eq $DevicesCondition -or $null -eq $DevicesCondition.DeviceFilter) {
        return "<span class='muted'>Not configured</span>"
    }

    $df = $DevicesCondition.DeviceFilter
    $lines = @()

    if ($df.Mode) { $lines += "Mode: $($df.Mode)" }
    if ($df.Rule) { $lines += "Rule: $($df.Rule)" }

    if (-not $lines) {
        return "<span class='muted'>Not configured</span>"
    }

    return ($lines | ForEach-Object { Encode-Html $_ }) -join "<br/>"
}

function Format-AuthFlows {
    param($Flows)

    if ($null -eq $Flows) {
        return "<span class='muted'>Not configured</span>"
    }

    $lines = @()
    $tm = $Flows.TransferMethods

    if ($tm) {
        if ($tm -is [string]) {
            $lines += "TransferMethods: $tm"
        } else {
            $lines += "TransferMethods: " + ($tm -join ", ")
        }
    }

    if (-not $lines) {
        return "<span class='muted'>Not configured</span>"
    }

    return ($lines | ForEach-Object { Encode-Html $_ }) -join "<br/>"
}

# ---------- Output path ----------
$timestamp        = Get-Date -Format "yyyy-MM-dd_HHmm"
$generatedDisplay = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$root             = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
$outFile          = Join-Path -Path $root -ChildPath "EntraID_CA_Policies_$timestamp.html"

# ---------- Tenant info ----------
$org = $null
try {
    $org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1 DisplayName, Id
} catch { }

$tenantName = if ($org.DisplayName) { $org.DisplayName } else { "Unknown tenant" }
$tenantId   = if ($org.Id) { $org.Id } else { "N/A" }

# ---------- Get CA policies ----------
$policies = Get-MgIdentityConditionalAccessPolicy -All

# ---------- HTML head ----------
$head = @"
<html>
<head>
<meta charset='UTF-8'>
<title>EntraID Conditional Access Policies - $timestamp</title>
<style>
    :root {
        --bg-main: #f3f4f6;
        --bg-header: #111827;
        --bg-header-accent: #1f2937;
        --bg-footer: #111827;
        --card-bg: #ffffff;
        --border-color: #d1d5db;
        --text-main: #111827;
        --text-muted: #4b5563;
        --text-invert: #f9fafb;
        --pill-enabled: #065f46;
        --pill-report: #92400e;
        --pill-disabled: #7f1d1d;
    }

    body {
        margin: 0;
        padding: 0;
        font-family: Segoe UI, Arial, sans-serif;
        font-size: 14px;
        background-color: var(--bg-main);
        color: var(--text-main);
    }

    .page-header {
        background: linear-gradient(90deg, var(--bg-header), var(--bg-header-accent));
        color: var(--text-invert);
        padding: 12px 20px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 1px solid #000;
    }

    .header-left h1 {
        margin: 0;
        font-size: 20px;
    }

    .header-left p,
    .header-right p {
        margin: 2px 0;
        font-size: 14px;
        color: #e5e7eb;
    }

    .header-right {
        text-align: right;
    }

    .main-content {
        padding: 16px 20px 40px 20px;
        max-width: 1400px;
        margin: 0 auto;
    }

    table {
        border-collapse: collapse;
        margin-top: 6px;
        margin-bottom: 12px;
        width: 100%;
        background-color: var(--card-bg);
    }

    th, td {
        border: 1px solid var(--border-color);
        padding: 4px 8px;
        vertical-align: top;
    }

    th {
        background-color: #f3f4f6;
        text-align: left;
    }

    .section-title {
        background-color: #e5e7eb;
        font-weight: bold;
        text-align: left;
    }

    .policy-index {
        margin-bottom: 12px;
    }

    .policy-index tr[data-policy-id] {
        cursor: pointer;
    }

    .policy-index tr[data-policy-id]:hover {
        background-color: #e5e7eb;
    }

    .pill {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 999px;
        font-size: 14px;      /* match table/body text size */
        font-weight: 600;
        color: #f9fafb;
    }

    .pill-enabled {
        background-color: var(--pill-enabled);
    }

    .pill-report {
        background-color: var(--pill-report);
    }

    .pill-disabled {
        background-color: var(--pill-disabled);
    }

    .muted {
        color: var(--text-muted);
        font-style: italic;
    }

    .toolbar {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        margin: 6px 0 8px 0;
        align-items: center;
    }

    .toolbar input {
        padding: 4px 8px;
        border-radius: 6px;
        border: 1px solid var(--border-color);
        font-size: 14px;
        min-width: 220px;
    }

    .toolbar button {
        padding: 4px 10px;
        border-radius: 6px;
        border: 1px solid var(--border-color);
        background-color: #ffffff;
        cursor: pointer;
        font-size: 14px;
    }

    .toolbar button:hover {
        background-color: #e5e7eb;
    }

    .info-banner {
        margin: 4px 0 14px 0;
        padding: 6px 10px;
        font-size: 14px;
        background-color: #e0f2fe;
        color: #0f172a;
        border: 1px solid #93c5fd;
        border-left-width: 4px;
        border-radius: 6px;
    }

    .policy-block {
        border: 1px solid var(--border-color);
        border-radius: 8px;
        margin-bottom: 16px;
        background-color: #ffffff;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }

    .policy-header {
        position: relative;
        padding: 8px 12px;
        padding-left: 26px; /* space for chevron */
        display: flex;
        justify-content: space-between;
        align-items: center;
        cursor: pointer;
        background-color: #f9fafb;
        border-bottom: 1px solid var(--border-color);
    }

    .toggle-icon {
        position: absolute;
        left: 8px;
        top: 50%;
        transform: translateY(-50%);
        font-size: 14px;
        line-height: 1;
        user-select: none;
        transition: transform 0.15s ease-out;
    }

    .policy-block.expanded .toggle-icon {
        transform: translateY(-50%) rotate(90deg);
    }

    .policy-header-main {
        display: flex;
        flex-direction: column;
    }

    .policy-title {
        font-weight: 600;
    }

    .policy-id {
        font-size: 10px;
        color: var(--text-muted);
    }

    .policy-content {
        padding: 6px 12px 10px 12px;
        display: none;
    }

    .policy-block.expanded .policy-content {
        display: block;
    }

    .page-footer {
        background-color: var(--bg-footer);
        color: var(--text-invert);
        text-align: center;
        padding: 8px 12px;
        font-size: 14px;
        border-top: 1px solid #000;
        position: sticky;
        bottom: 0;
    }
</style>
</head>
<body>
<div class="page-header">
  <div class="header-left">
    <h1>EntraID Conditional Access Policies</h1>
    <p>Tenant: $tenantName</p>
    <p>Tenant ID: $tenantId</p>
  </div>
  <div class="header-right">
    <p>Generated: $generatedDisplay</p>
    <p>Report file: EntraID_CA_Policies_$timestamp.html</p>
  </div>
</div>

<div class="main-content">
  <h2>Policy Index</h2>
  <table class="policy-index">
    <tr>
      <th>Name</th>
      <th>State</th>
      <th>Policy Id</th>
    </tr>
"@

# ---------- Index rows ----------
$indexRows = foreach ($p in ($policies | Sort-Object DisplayName)) {
    $state = $p.State
    $pillClass = switch ($state) {
        "enabled"                          { "pill pill-enabled" }
        "enabledForReportingButNotEnforced" { "pill pill-report" }
        "disabled"                         { "pill pill-disabled" }
        default                            { "pill pill-report" }
    }

@"
    <tr data-policy-id="$($p.Id)">
      <td>$(Encode-Html $p.DisplayName)</td>
      <td><span class="$pillClass">$state</span></td>
      <td>$(Encode-Html $p.Id)</td>
    </tr>
"@
}

$indexClose = @"
  </table>
  <div class="toolbar">
    <input type="text" id="policySearch" placeholder="Filter policies by name..." oninput="filterPolicies()" />
    <button type="button" onclick="expandAll()">Expand all</button>
    <button type="button" onclick="collapseAll()">Collapse all</button>
  </div>

  <div class="info-banner">
    Tip: Click a policy in the index to jump to it. Use the arrow icon or the policy header to expand or collapse details.
    You can also use the "Expand all" and "Collapse all" buttons for bulk actions.
  </div>
"@

# ---------- Policy detail blocks ----------
$body = foreach ($p in ($policies | Sort-Object DisplayName)) {
    $state = $p.State
    $pillClass = switch ($state) {
        "enabled"                          { "pill pill-enabled" }
        "enabledForReportingButNotEnforced" { "pill pill-report" }
        "disabled"                         { "pill pill-disabled" }
        default                            { "pill pill-report" }
    }

    $users      = $p.Conditions.Users
    $apps       = $p.Conditions.Applications
    $locations  = $p.Conditions.Locations
    $clientApps = $p.Conditions.ClientAppTypes
    $platforms  = $p.Conditions.Platforms
    $authFlows  = $p.Conditions.AuthenticationFlows
    $grant      = $p.GrantControls
    $session    = $p.SessionControls

@"
<div class="policy-block collapsed" id="policy-$($p.Id)">
  <div class="policy-header" onclick="togglePolicy('policy-$($p.Id)')">
    <span class="toggle-icon">&#9654;</span>
    <div class="policy-header-main">
      <span class="policy-title">$(Encode-Html $p.DisplayName)</span>
      <span class="policy-id">Id: $(Encode-Html $p.Id)</span>
    </div>
    <span class="$pillClass">$state</span>
  </div>
  <div class="policy-content">

<table>
<tr><th class='section-title' colspan='2'>Summary</th></tr>
<tr><th>Description</th><td>$(Encode-Html $p.Description)</td></tr>
<tr><th>Created</th><td>$($p.CreatedDateTime)</td></tr>
<tr><th>Modified</th><td>$($p.ModifiedDateTime)</td></tr>
</table>

<table>
<tr><th class='section-title' colspan='2'>Assignments - Users</th></tr>
<tr><th>Include Users</th><td>$(Format-List $users.IncludeUsers)</td></tr>
<tr><th>Exclude Users</th><td>$(Format-List $users.ExcludeUsers)</td></tr>
<tr><th>Include Groups</th><td>$(Format-List $users.IncludeGroups)</td></tr>
<tr><th>Exclude Groups</th><td>$(Format-List $users.ExcludeGroups)</td></tr>
<tr><th>Include Roles</th><td>$(Format-List $users.IncludeRoles)</td></tr>
<tr><th>Exclude Roles</th><td>$(Format-List $users.ExcludeRoles)</td></tr>
</table>

<table>
<tr><th class='section-title' colspan='2'>Assignments - Cloud apps &amp; Actions</th></tr>
<tr><th>Include Apps</th><td>$(Format-List $apps.IncludeApplications)</td></tr>
<tr><th>Exclude Apps</th><td>$(Format-List $apps.ExcludeApplications)</td></tr>
<tr><th>Include User Actions</th><td>$(Format-List $apps.IncludeUserActions)</td></tr>
<tr><th>Auth Context IDs</th><td>$(Format-List $apps.IncludeAuthenticationContextClassReferences)</td></tr>
</table>

<table>
<tr><th class='section-title' colspan='2'>Conditions</th></tr>
<tr><th>Client App Types</th><td>$(Format-List $clientApps)</td></tr>
<tr><th>Include Locations</th><td>$(Format-List $locations.IncludeLocations)</td></tr>
<tr><th>Exclude Locations</th><td>$(Format-List $locations.ExcludeLocations)</td></tr>
<tr><th>Platforms</th><td>$(Format-Platforms $platforms)</td></tr>
<tr><th>Device Filter</th><td>$(Format-DeviceFilter $p.Conditions.Devices)</td></tr>
<tr><th>Sign-in Risk</th><td>$(Format-List $p.Conditions.SignInRiskLevels)</td></tr>
<tr><th>User Risk</th><td>$(Format-List $p.Conditions.UserRiskLevels)</td></tr>
<tr><th>Authentication Flows</th><td>$(Format-AuthFlows $authFlows)</td></tr>
</table>

<table>
<tr><th class='section-title' colspan='2'>Access Controls - Grant</th></tr>
<tr><th>Built-in Controls</th><td>$(Format-List $grant.BuiltInControls)</td></tr>
<tr><th>Custom Auth Factors</th><td>$(Format-List $grant.CustomAuthenticationFactors)</td></tr>
<tr><th>Terms of Use</th><td>$(Format-List $grant.TermsOfUse)</td></tr>
<tr><th>Operator</th><td>$(Encode-Html $grant.Operator)</td></tr>
<tr><th>Authentication Strength Id</th><td>$(Encode-Html $grant.AuthenticationStrength.Id)</td></tr>
</table>

<table>
<tr><th class='section-title' colspan='2'>Access Controls - Session</th></tr>
<tr><th>Application Enforced Restrictions</th><td>$(Format-SessionObject $session.ApplicationEnforcedRestrictions)</td></tr>
<tr><th>Cloud App Security</th><td>$(Format-SessionObject $session.CloudAppSecurity)</td></tr>
<tr><th>Persistent Browser</th><td>$(Format-SessionObject $session.PersistentBrowser)</td></tr>
<tr><th>Sign-in Frequency</th><td>$(Format-SessionObject $session.SignInFrequency)</td></tr>
</table>

  </div>
</div>
"@
}

# ---------- Footer + JS ----------
$footer = @"
</div> <!-- /main-content -->

<div class="page-footer">
  <p>Conditional Access snapshot generated by My IT Solutions export script on $generatedDisplay</p>
</div>

<script>
document.addEventListener('DOMContentLoaded', function () {
  // Click index row to scroll to policy
  var rows = document.querySelectorAll('.policy-index tr[data-policy-id]');
  rows.forEach(function (row) {
    row.addEventListener('click', function () {
      var id = row.getAttribute('data-policy-id');
      var block = document.getElementById('policy-' + id);
      if (block) {
        block.classList.add('expanded');
        var y = block.getBoundingClientRect().top + window.pageYOffset - 80;
        window.scrollTo({ top: y, behavior: 'smooth' });
      }
    });
  });
});

function togglePolicy(blockId) {
  var block = document.getElementById(blockId);
  if (!block) return;
  block.classList.toggle('expanded');
}

function expandAll() {
  document.querySelectorAll('.policy-block').forEach(function (block) {
    block.classList.add('expanded');
    block.style.display = '';
  });
}

function collapseAll() {
  document.querySelectorAll('.policy-block').forEach(function (block) {
    block.classList.remove('expanded');
    block.style.display = '';
  });
}

function filterPolicies() {
  var term = document.getElementById('policySearch').value.toLowerCase();
  var rows = document.querySelectorAll('.policy-index tr[data-policy-id]');

  rows.forEach(function (row) {
    var nameCell = row.querySelector('td');
    var id = row.getAttribute('data-policy-id');
    var block = document.getElementById('policy-' + id);

    if (!nameCell) return;

    var name = nameCell.innerText.toLowerCase();
    var match = name.indexOf(term) !== -1 || term === '';

    row.style.display = match ? '' : 'none';
    if (block) {
      block.style.display = match ? '' : 'none';
    }
  });
}
</script>

</body>
</html>
"@

# ---------- Write file ----------
$fullHtml = $head + ($indexRows -join "`r`n") + $indexClose + ($body -join "`r`n") + $footer
$fullHtml | Out-File -FilePath $outFile -Encoding UTF8

Write-Host "Exported CA policy report to: $outFile"
