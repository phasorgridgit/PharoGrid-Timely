# PhasorGrid Timely — starter data seeding (PnP PowerShell)
# Run AFTER provision-sharepoint.ps1 has created the lists/columns.
#
# Usage (note the -ClientId — Connect-PnPOnline -Interactive alone will fail
# with "Please specify a valid client id" on modern tenants, since Microsoft
# retired the shared "PnP Management Shell" app PnP used to fall back on):
#   Connect-PnPOnline -Url "https://phasorgrid.sharepoint.com/sites/WorkHub" `
#     -Interactive -ClientId "<same WORKHUB_CLIENT_ID from your .env>"
#   .\seed-sharepoint-data.ps1
#
# Idempotent: every Add-Item call below looks the row up by Title first and
# skips it if it's already there, so re-running this after adding more rows
# to the arrays at the top won't duplicate anything.
#
# What this seeds:
#   - WorkHub_Clients   (2 starter clients)
#   - WorkHub_Projects  (2 starter projects, linked to the clients above)
#   - WorkHub_TeamMembers (roster — edit the array below to match your org)
#   - WorkHub_TimeOffBalances (one row per teammate, resolving their real
#     Entra object id via Get-PnPAzureADUser — the same id value the app
#     reads back from Microsoft Graph `/me`, since EmployeeId is a Text
#     column, not a Person column — see provision-sharepoint.ps1 for why)
#
# Edit the arrays below to match your real clients/projects/team before
# running this against a production site — what's here is only a
# reasonable starting point mirroring the app's built-in mock data.

# ---- Preflight: fail loudly instead of silently no-op'ing every call ------
# PowerShell doesn't stop on a failed cmdlet by default (Add-PnPListItem
# just writes a non-terminating error and moves on), so without this check
# a "not signed in" session runs the WHOLE script, prints a misleading
# "added ..." line for every row, and creates nothing at all in SharePoint.
try {
  $conn = Get-PnPConnection -ErrorAction Stop
  if (-not $conn) { throw "no active connection" }
  Write-Host "Connected as $($conn.PSCredential.UserName) to $($conn.Url)" -ForegroundColor Green
} catch {
  Write-Error "Not connected to SharePoint. Run Connect-PnPOnline first (see the Usage comment at the top of this script), then re-run this script."
  exit 1
}

$script:failureCount = 0

function Add-ItemIfNotExists {
  param(
    [string]$ListName,
    [string]$Title,
    [hashtable]$Fields
  )
  try {
    $existing = Get-PnPListItem -List $ListName -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$Title</Value></Eq></Where></Query></View>" -ErrorAction Stop
  } catch {
    Write-Error "  Couldn't read $ListName — is the list name right, and did provision-sharepoint.ps1 run first? $($_.Exception.Message)"
    $script:failureCount++
    return $null
  }

  if ($existing) {
    Write-Host "  '$Title' already exists in $ListName, skipping"
    return $existing[0]
  }

  try {
    $allFields = @{ Title = $Title } + $Fields
    $item = Add-PnPListItem -List $ListName -Values $allFields -ErrorAction Stop
    Write-Host "  added '$Title' to $ListName (id $($item.Id))" -ForegroundColor Green
    return $item
  } catch {
    Write-Error "  FAILED to add '$Title' to $ListName — $($_.Exception.Message)"
    $script:failureCount++
    return $null
  }
}

# Resolves a UPN/email to the same Entra object-id string the app gets back
# from Microsoft Graph `/me` — this is what goes in every EmployeeId text
# column so filters like `fields/EmployeeId eq '<id>'` actually match what
# the running app writes.
function Get-EmployeeObjectId {
  param([string]$Upn)
  try {
    return (Get-PnPAzureADUser -Identity $Upn -ErrorAction Stop).Id
  } catch {
    Write-Warning "Couldn't resolve '$Upn' via Get-PnPAzureADUser — check the email is correct and that your connected account can read the tenant's users. Skipping their balance row."
    return $null
  }
}

# ---- Clients --------------------------------------------------------------
Write-Host "Seeding WorkHub_Clients..."
$clients = @(
  @{ Title = "Meralco"; Fields = @{ DefaultRate = 145 } },
  @{ Title = "Davao Light & Power"; Fields = @{ DefaultRate = 110 } }
)
$clientItems = @{}
foreach ($c in $clients) {
  $item = Add-ItemIfNotExists -ListName "WorkHub_Clients" -Title $c.Title -Fields $c.Fields
  if ($item) { $clientItems[$c.Title] = $item }
}

# ---- Projects (linked to clients above by name, not a lookup id) ----------
Write-Host "Seeding WorkHub_Projects..."
$projects = @(
  @{ Title = "Grid Modernization"; Client = "Meralco"; Rate = 145 },
  @{ Title = "Meter Rollout Ph.2"; Client = "Davao Light & Power"; Rate = 110 }
)
foreach ($p in $projects) {
  if (-not $clientItems.ContainsKey($p.Client)) {
    Write-Error "  Skipping project '$($p.Title)' — its client '$($p.Client)' wasn't created above."
    $script:failureCount++
    continue
  }
  Add-ItemIfNotExists -ListName "WorkHub_Projects" -Title $p.Title -Fields @{
    ClientName  = $p.Client
    HourlyRate  = $p.Rate
    Status      = "Active"
  } | Out-Null
}

# ---- Team roster ------------------------------------------------------------
# Edit this list to match your real team — Email should be each person's UPN
# so Get-EmployeeObjectId (below, for balances) can resolve it.
Write-Host "Seeding WorkHub_TeamMembers..."
$team = @(
  @{ Title = "Yash";     Email = "yash@phasorgrid.com";     Role = "Admin"; Group = "Protection Engineering"; Rate = 120 },
  @{ Title = "Dinesh K"; Email = "dinesh@phasorgrid.com";   Role = "Admin"; Group = "Protection Engineering"; Rate = 110 },
  @{ Title = "Mahesh N"; Email = "mahesh@phasorgrid.com";   Role = "Admin"; Group = "Compliance";             Rate = 110 },
  @{ Title = "Santhosh K"; Email = "santhosh@phasorgrid.com"; Role = "Member"; Group = "Compliance";          Rate = 95 },
  @{ Title = "Sneha R";  Email = "sneha@phasorgrid.com";    Role = "Member"; Group = "QA";                    Rate = 90 },
  @{ Title = "Venkata P"; Email = "venkata@phasorgrid.com"; Role = "Owner"; Group = "Leadership";             Rate = 150 }
  @{ Title = "Phani Sadhanala"; Email = "phani@phasorgrid.com"; Role = "Admin"; Group = "sde";             Rate = 90}
)
foreach ($t in $team) {
  Add-ItemIfNotExists -ListName "WorkHub_TeamMembers" -Title $t.Title -Fields @{
    Email        = $t.Email
    Role         = $t.Role
    Group        = $t.Group
    BillableRate = $t.Rate
  } | Out-Null
}

# ---- Time-off balances (one row per teammate, real Entra object ids) -------
Write-Host "Seeding WorkHub_TimeOffBalances..."
foreach ($t in $team) {
  $objectId = Get-EmployeeObjectId -Upn $t.Email
  if (-not $objectId) { continue }
  Add-ItemIfNotExists -ListName "WorkHub_TimeOffBalances" -Title "$($t.Title) — balance" -Fields @{
    EmployeeId    = $objectId
    AccruedDays   = 20
    UsedDays      = 0
    AvailableDays = 20
  } | Out-Null
}

Write-Host ""
if ($script:failureCount -gt 0) {
  Write-Warning "Finished with $($script:failureCount) failure(s) above — those rows were NOT created. Fix the issue and re-run; existing rows are left alone."
  exit 1
} else {
  Write-Host "Done seeding — every row above either already existed or was created successfully." -ForegroundColor Green
  Write-Host "Adjust the arrays at the top of this script and re-run any time."
}