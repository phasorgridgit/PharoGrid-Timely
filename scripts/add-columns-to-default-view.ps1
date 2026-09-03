# PhasorGrid Timely — fix "data isn't showing up in SharePoint" (PnP PowerShell)
#
# WHAT THIS FIXES:
# Add-PnPField creates a column in the list schema, but does NOT add it to
# the "All Items" view unless you pass -AddToDefaultView. That means the app
# can read/write every field fine over Microsoft Graph — Timesheets,
# invoicing, etc. all work — but opening the list directly in a browser
# shows only "Title", because the other columns were never added to the
# view. It looks exactly like data isn't syncing, even though it is.
#
# provision-sharepoint.ps1 now creates new fields with -AddToDefaultView,
# but that only applies going forward — it can't retrofit fields an earlier
# run of the script already created (Add-PnPField skips existing fields).
# This script fixes that for a site that's already provisioned: it only
# changes which columns the "All Items" view displays. It does NOT touch
# any data, delete anything, or change field types.
#
# Usage:
#   Connect-PnPOnline -Url "https://phasorgrid.sharepoint.com/sites/WorkHub" -Interactive
#   .\add-columns-to-default-view.ps1
#
# Safe to re-run.

try {
  $conn = Get-PnPConnection -ErrorAction Stop
  if (-not $conn) { throw "no active connection" }
  Write-Host "Connected as $($conn.PSCredential.UserName) to $($conn.Url)" -ForegroundColor Green
} catch {
  Write-Error "Not connected to SharePoint. Run Connect-PnPOnline first, then re-run this script."
  exit 1
}

# Only adds fields that actually exist on the list right now — if your site
# still has old approval-era fields (ApproverId, SubmittedOn, etc.) or is
# missing a newer field (BillableHours, NonBillableHours) because it hasn't
# been re-provisioned, this just works with whatever's really there instead
# of erroring out.
function Set-WorkHubDefaultView {
  param(
    [Parameter(Mandatory)][string]$List,
    [Parameter(Mandatory)][string[]]$PreferredFields
  )
  try {
    $existing = (Get-PnPField -List $List -ErrorAction Stop) | Select-Object -ExpandProperty InternalName
  } catch {
    Write-Warning "  Couldn't read fields for $List — does the list exist? $($_.Exception.Message)"
    return
  }
  $fields = $PreferredFields | Where-Object { $existing -contains $_ }
  if ($fields.Count -eq 0) {
    Write-Warning "  None of the expected fields exist on $List yet — skipping."
    return
  }
  try {
    Set-PnPView -List $List -Identity "All Items" -Fields $fields -ErrorAction Stop
    Write-Host "  $List -> $($fields -join ', ')" -ForegroundColor Green
  } catch {
    Write-Warning "  Couldn't update the default view for $List — $($_.Exception.Message)"
  }
}

Write-Host "Updating default views..."

Set-WorkHubDefaultView -List "WorkHub_Clients" -PreferredFields @(
  "Title", "ContactName", "ContactEmail", "DefaultRate", "BillingAddress"
)

Set-WorkHubDefaultView -List "WorkHub_Projects" -PreferredFields @(
  "Title", "ClientName", "HourlyRate", "Budget", "BudgetHours", "Status", "TeamMembers",
  # only shown if this site still has the pre-migration lookup column:
  "ClientId"
)

Set-WorkHubDefaultView -List "WorkHub_Timesheets" -PreferredFields @(
  "Title", "EmployeeName", "WeekStart", "WeekEnd", "TotalHours", "BillableHours", "NonBillableHours",
  # kept only so old sites that still have these approval/id-era columns can
  # see them too, until they're re-provisioned onto the current schema:
  "EmployeeId", "Status", "ApproverId", "SubmittedOn", "ApprovedOn"
)

Set-WorkHubDefaultView -List "WorkHub_TimeEntries" -PreferredFields @(
  "Title", "EmployeeName", "ProjectName", "ClientName", "TaskName", "Date", "StartTime", "EndTime", "Hours", "Billable", "Description", "Tags",
  # only shown if this site still has these pre-migration columns (old
  # id-based/timer model, or the old approval workflow):
  "EmployeeId", "ProjectId", "ClientId", "DurationSeconds", "Source", "Status", "TimesheetWeekId"
)

Set-WorkHubDefaultView -List "WorkHub_Invoices" -PreferredFields @(
  "Title", "ClientName", "PeriodStart", "PeriodEnd", "Amount", "Status", "DueDate", "FileUrl",
  # only shown if this site still has the pre-migration lookup column:
  "ClientId"
)

Set-WorkHubDefaultView -List "WorkHub_SyncLog" -PreferredFields @(
  "Title", "Timestamp", "ActionType", "Status", "Detail"
)

Set-WorkHubDefaultView -List "WorkHub_Settings" -PreferredFields @(
  "Title", "RetentionYears",
  "DefaultApproverId" # only shown if your site still has the old column
)

Set-WorkHubDefaultView -List "WorkHub_TimeOff" -PreferredFields @(
  "Title", "EmployeeId", "Type", "StartDate", "EndDate", "Days", "Notes", "Status"
)

Set-WorkHubDefaultView -List "WorkHub_TimeOffBalances" -PreferredFields @(
  "Title", "EmployeeId", "AccruedDays", "UsedDays", "AvailableDays"
)

Set-WorkHubDefaultView -List "WorkHub_TeamMembers" -PreferredFields @(
  "Title", "Email", "Role", "Group", "BillableRate"
)

Write-Host ""
Write-Host "Done. Refresh each list in the browser — you should now see every column, not just Title." -ForegroundColor Green