# PhasorGrid Timely — OPTIONAL cleanup script (PnP PowerShell)
#
# provision-sharepoint.ps1 no longer creates any of the columns below — the
# app now references employees/projects/clients by NAME and logs time as
# Hours directly, with no timer and no approval workflow. This script is for
# a site that was provisioned BEFORE that change and still has the old
# columns sitting around unused:
#   WorkHub_TimeEntries: EmployeeId, ProjectId, ClientId,
#                         DurationSeconds, Source, Status, TimesheetWeekId
#   WorkHub_Projects:    ClientId
#   WorkHub_Invoices:    ClientId
#   WorkHub_Timesheets:  EmployeeId, Status, ApproverId, SubmittedOn, ApprovedOn
#
# You do NOT have to run this — those columns are harmless if left in place,
# the app just no longer reads or writes any of them. Run this only if you
# want them gone from the SharePoint UI.
#
# This script is destructive (it deletes columns, which deletes whatever
# data was stored in them for every existing item) and requires explicit
# confirmation before touching anything. It does NOT touch the site, the
# lists themselves, or any other columns.
#
# Usage:
#   Connect-PnPOnline -Url "https://phasorgrid.sharepoint.com/sites/WorkHub" -Interactive
#   .\cleanup-obsolete-fields.ps1
#   .\cleanup-obsolete-fields.ps1 -WhatIf   # preview only, deletes nothing

param(
  [switch]$WhatIf
)

try {
  $conn = Get-PnPConnection -ErrorAction Stop
  if (-not $conn) { throw "no active connection" }
  Write-Host "Connected as $($conn.PSCredential.UserName) to $($conn.Url)" -ForegroundColor Green
} catch {
  Write-Error "Not connected to SharePoint. Run Connect-PnPOnline first, then re-run this script."
  exit 1
}

$targets = @(
  @{ List = "WorkHub_TimeEntries"; Fields = @("EmployeeId", "ProjectId", "ClientId", "DurationSeconds", "Source", "Status", "TimesheetWeekId") },
  @{ List = "WorkHub_Projects";    Fields = @("ClientId") },
  @{ List = "WorkHub_Invoices";    Fields = @("ClientId") },
  @{ List = "WorkHub_Timesheets";  Fields = @("EmployeeId", "Status", "ApproverId", "SubmittedOn", "ApprovedOn") }
)

Write-Host ""
Write-Host "The following columns will be PERMANENTLY DELETED, along with any data in them:" -ForegroundColor Yellow
foreach ($t in $targets) {
  foreach ($f in $t.Fields) { Write-Host "  - $($t.List).$f" }
}
Write-Host ""
Write-Host "Only columns that actually exist on your site will be touched — anything already migrated is skipped." -ForegroundColor Yellow
Write-Host ""

if ($WhatIf) {
  Write-Host "-WhatIf passed — nothing will be deleted. Re-run without -WhatIf to apply." -ForegroundColor Cyan
  exit 0
}

$confirm = Read-Host "Type DELETE to confirm you want to permanently remove these columns"
if ($confirm -ne "DELETE") {
  Write-Host "Aborted — nothing was changed."
  exit 0
}

foreach ($t in $targets) {
  foreach ($f in $t.Fields) {
    try {
      Get-PnPField -List $t.List -Identity $f -ErrorAction Stop | Out-Null
      Remove-PnPField -List $t.List -Identity $f -Force -ErrorAction Stop
      Write-Host "  removed $($t.List).$f" -ForegroundColor Green
    } catch {
      Write-Host "  $($t.List).$f already gone, skipping" -ForegroundColor DarkGray
    }
  }
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
