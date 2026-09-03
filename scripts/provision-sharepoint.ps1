# PhasorGrid Timely — SharePoint provisioning script (PnP PowerShell)
# Run once per environment (Dev / Prod) after the "WorkHub" site itself exists.
# Requires: Install-Module PnP.PowerShell
#
# Usage:
#   Connect-PnPOnline -Url "https://phasorgrid.sharepoint.com/sites/WorkHub" -Interactive
#   .\provision-sharepoint.ps1
#
# Safe to re-run: -ErrorAction SilentlyContinue on list/field creation means
# it just skips anything that already exists.
#
# IMPORTANT — every Add-PnPField call below passes -AddToDefaultView. The
# one UserMulti field (TeamMembers) is the exception: Add-PnPFieldFromXml
# doesn't accept that parameter, so it's added to the default view as a
# separate step instead — see Add-WorkHubUserMultiField below. Without
# this, Add-PnPField creates the column in the list schema — the app can
# read/write it over Graph just fine — but it never shows up in the "All
# Items" view in a browser, which makes it look like data isn't syncing
# when it actually is. Because -ErrorAction SilentlyContinue skips fields
# that already exist, re-running this script will NOT retrofit
# -AddToDefaultView onto fields an earlier run already created — see
# scripts/add-columns-to-default-view.ps1 for a small, non-destructive
# script that fixes that on an already-provisioned site.
#
# IMPORTANT — field types below are deliberately NOT "Lookup"/"User"/"URL"
# for the columns the app writes to directly:
#   - EmployeeName / ProjectName / ClientName are plain Text, not
#     Lookup/Person columns. Time entries reference the employee, project,
#     and client by NAME — never by SharePoint item id, Entra object id, or
#     any other id. This is a deliberate design choice, not a shortcut: it
#     keeps every WorkHub_TimeEntries row human-readable on its own, with no
#     joins needed to see who logged what, on what project, for which
#     client.
#   - FileUrl (on WorkHub_Invoices) is Text, not URL/Hyperlink. The app just
#     stores the plain webUrl string OneDrive/SharePoint returns after an
#     upload — a real Hyperlink column expects a {Url, Description} object,
#     which is a different write shape than the app produces.
#   - TeamMembers on WorkHub_Projects is the one exception: it's set by
#     admins directly in the SharePoint UI, never written by the app, so a
#     real UserMulti (Person) column is fine there.
#
# NOTE: There is no approval workflow and no timer in this app. Employees
# type their hours directly into the Timesheets page; every entry is
# written to WorkHub_TimeEntries immediately as EmployeeName / ProjectName /
# ClientName / Date / Hours / Billable / Description / Tags. Hours is
# always a plain decimal (8, 7.5, 2.25) — nothing here is ever stored or
# displayed in seconds, and there is no StartTime/EndTime/DurationSeconds.
# WorkHub_Timesheets, if used, is only a weekly summary/cache — the
# Timesheets page itself always reads live from WorkHub_TimeEntries.
#
# NOTE on UserMulti fields: current PnP PowerShell's Add-PnPField -Type
# parameter only takes the SharePoint FieldType enum, which has no
# UserMulti value, so that one field is created with Add-PnPFieldFromXml
# instead (a raw CAML field definition) — see the helper function right
# below this comment block. Every other field (Text/Choice/Number/etc.) is
# unaffected and still uses plain Add-PnPField.
#
# See scripts/provision-sharepoint-graph.mjs for the Graph-API equivalent of
# this same script, and src/graph/*.js for exactly how each field gets read
# and written.

# Creates a multi-value Person column via CAML XML — Add-PnPField's -Type
# enum has no UserMulti value. TeamMembers (on WorkHub_Projects) is the only
# field that needs this; every other employee/project/client reference in
# this app is a plain Name text field, not a Lookup or a Person column.
function Add-WorkHubUserMultiField {
  param(
    [Parameter(Mandatory)][string]$List,
    [Parameter(Mandatory)][string]$InternalName,
    [Parameter(Mandatory)][string]$DisplayName
  )
  $fieldXml = "<Field Type=""UserMulti"" DisplayName=""$DisplayName"" Name=""$InternalName"" StaticName=""$InternalName"" Mult=""TRUE"" UserSelectionMode=""PeopleOnly"" UserSelectionScope=""0"" />"
  # NOTE: Add-PnPFieldFromXml does NOT accept -AddToDefaultView in current
  # PnP.PowerShell versions (only Add-PnPField does) — passing it throws
  # "A parameter cannot be found that matches parameter name
  # 'AddToDefaultView'" and skips creating this field entirely, which is
  # why WorkHub_Projects.TeamMembers may be missing after an earlier run of
  # this script. Create the field first, then add it to the default view
  # as a separate step; wrapped in try/catch since the exact cmdlet name
  # for that step has changed across PnP.PowerShell versions too, and this
  # field being missing from the default view is harmless (the app never
  # reads/writes it — TeamMembers is admin-only, set by hand in the UI).
  Add-PnPFieldFromXml -List $List -FieldXml $fieldXml -ErrorAction SilentlyContinue
  try { Add-PnPFieldToDefaultView -List $List -Field $InternalName -ErrorAction Stop } catch { }
}

Write-Host "Creating document libraries..."
New-PnPList -Title "WorkHub Invoices" -Template DocumentLibrary -ErrorAction SilentlyContinue
New-PnPList -Title "WorkHub Documents" -Template DocumentLibrary -ErrorAction SilentlyContinue

Write-Host "Creating WorkHub_Clients..."
New-PnPList -Title "WorkHub_Clients" -Template GenericList -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Clients" -DisplayName "ContactName" -InternalName "ContactName" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Clients" -DisplayName "ContactEmail" -InternalName "ContactEmail" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Clients" -DisplayName "DefaultRate" -InternalName "DefaultRate" -Type Currency -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Clients" -DisplayName "BillingAddress" -InternalName "BillingAddress" -Type Note -AddToDefaultView -ErrorAction SilentlyContinue

Write-Host "Creating WorkHub_Projects..."
New-PnPList -Title "WorkHub_Projects" -Template GenericList -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Projects" -DisplayName "ClientName" -InternalName "ClientName" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Projects" -DisplayName "HourlyRate" -InternalName "HourlyRate" -Type Currency -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Projects" -DisplayName "Budget" -InternalName "Budget" -Type Currency -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Projects" -DisplayName "BudgetHours" -InternalName "BudgetHours" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Projects" -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Active","On hold","Completed" -AddToDefaultView -ErrorAction SilentlyContinue
Add-WorkHubUserMultiField -List "WorkHub_Projects" -InternalName "TeamMembers" -DisplayName "TeamMembers"

# WorkHub_Timesheets is a weekly summary/cache only — there is no approval
# workflow. Employees record their own time directly into
# WorkHub_TimeEntries (the source of truth); this list, if used at all, just
# holds a rolled-up total per employee per week for faster reporting.
Write-Host "Creating WorkHub_Timesheets..."
New-PnPList -Title "WorkHub_Timesheets" -Template GenericList -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Timesheets" -DisplayName "EmployeeName" -InternalName "EmployeeName" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Timesheets" -DisplayName "WeekStart" -InternalName "WeekStart" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Timesheets" -DisplayName "WeekEnd" -InternalName "WeekEnd" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Timesheets" -DisplayName "TotalHours" -InternalName "TotalHours" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Timesheets" -DisplayName "BillableHours" -InternalName "BillableHours" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Timesheets" -DisplayName "NonBillableHours" -InternalName "NonBillableHours" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Set-PnPField -List "WorkHub_Timesheets" -Identity "EmployeeName" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "Creating WorkHub_TimeEntries..."
New-PnPList -Title "WorkHub_TimeEntries" -Template GenericList -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TimeEntries" -DisplayName "EmployeeName" -InternalName "EmployeeName" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TimeEntries" -DisplayName "ProjectName" -InternalName "ProjectName" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TimeEntries" -DisplayName "ClientName" -InternalName "ClientName" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
# TaskName is plain text, same NAME-based convention as ProjectName — which
# specific task (from WorkHub_ProjectTasks) under this entry's project the
# employee was working on, or "" when no specific task was picked. Never a
# Lookup id back to WorkHub_ProjectTasks.
Add-PnPField -List "WorkHub_TimeEntries" -DisplayName "TaskName" -InternalName "TaskName" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TimeEntries" -DisplayName "Date" -InternalName "Date" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
# Hours is a plain decimal number — 8, 7.5, 2.25 — entered directly by the
# employee. There is no timer, no clock-in/clock-out, and nothing here is
# ever stored or displayed in seconds.
Add-PnPField -List "WorkHub_TimeEntries" -DisplayName "Hours" -InternalName "Hours" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
# StartTime/EndTime are plain "HH:mm" TEXT, not a DateTime column and not a
# revival of the old timer — there is still no running clock and nothing is
# captured automatically. These are just two optional fields the employee
# types by hand (e.g. "09:00" / "17:30") so the Calendar page can draw a
# real time-of-day block instead of just a flat daily total. Text avoids
# any timezone conversion entirely; Date already carries which day it was.
Add-PnPField -List "WorkHub_TimeEntries" -DisplayName "StartTime" -InternalName "StartTime" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TimeEntries" -DisplayName "EndTime" -InternalName "EndTime" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TimeEntries" -DisplayName "Billable" -InternalName "Billable" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TimeEntries" -DisplayName "Description" -InternalName "Description" -Type Note -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TimeEntries" -DisplayName "Tags" -InternalName "Tags" -Type MultiChoice -Choices "scada","qa","call","internal","analysis" -AddToDefaultView -ErrorAction SilentlyContinue
# Index the columns the app actually filters on: EmployeeName (every read is
# scoped to one employee) and Date (every read is scoped to a date range).
# There is no Status field — with no timer and no approval step, every row
# in this list is simply a logged, finished entry; a status distinguishing
# "open" from "completed" or "submitted" from "approved" has nothing left to
# describe, so it was removed rather than kept unused.
Set-PnPField -List "WorkHub_TimeEntries" -Identity "EmployeeName" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List "WorkHub_TimeEntries" -Identity "Date" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "Creating WorkHub_ProjectTasks..."
New-PnPList -Title "WorkHub_ProjectTasks" -Template GenericList -ErrorAction SilentlyContinue
# ProjectName is plain Text here too, same convention as ProjectName on
# WorkHub_TimeEntries — tasks reference the project by NAME, never by
# SharePoint item id. The task's own text lives in the list's built-in
# Title column, so no separate column is needed for it.
Add-PnPField -List "WorkHub_ProjectTasks" -DisplayName "ProjectName" -InternalName "ProjectName" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
Set-PnPField -List "WorkHub_ProjectTasks" -Identity "ProjectName" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "Creating WorkHub_Invoices..."
New-PnPList -Title "WorkHub_Invoices" -Template GenericList -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Invoices" -DisplayName "ClientName" -InternalName "ClientName" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Invoices" -DisplayName "PeriodStart" -InternalName "PeriodStart" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Invoices" -DisplayName "PeriodEnd" -InternalName "PeriodEnd" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Invoices" -DisplayName "Amount" -InternalName "Amount" -Type Currency -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Invoices" -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Draft","Sent","Paid","Overdue" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Invoices" -DisplayName "DueDate" -InternalName "DueDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Invoices" -DisplayName "FileUrl" -InternalName "FileUrl" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Invoices" -DisplayName "LineItemsJson" -InternalName "LineItemsJson" -Type Note -AddToDefaultView -ErrorAction SilentlyContinue

Write-Host "Creating WorkHub_SyncLog..."
New-PnPList -Title "WorkHub_SyncLog" -Template GenericList -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_SyncLog" -DisplayName "Timestamp" -InternalName "Timestamp" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_SyncLog" -DisplayName "ActionType" -InternalName "ActionType" -Type Choice -Choices "Create","Update","Delete","Sync","Error" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_SyncLog" -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Success","Failed" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_SyncLog" -DisplayName "Detail" -InternalName "Detail" -Type Note -AddToDefaultView -ErrorAction SilentlyContinue

Write-Host "Creating WorkHub_Settings (single item, holds retention config)..."
New-PnPList -Title "WorkHub_Settings" -Template GenericList -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_Settings" -DisplayName "RetentionYears" -InternalName "RetentionYears" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue

Write-Host "Creating WorkHub_TimeOff..."
New-PnPList -Title "WorkHub_TimeOff" -Template GenericList -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TimeOff" -DisplayName "EmployeeId" -InternalName "EmployeeId" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
# Indexed because getMyRequests() filters on it (fields/EmployeeId eq
# '...') — every other column any query filters on in this app
# (WorkHub_TimeEntries.EmployeeName/Date, WorkHub_TimeOffBalances.
# EmployeeId, WorkHub_ProjectTasks.ProjectName) is indexed the same way.
# Left un-indexed, this filter works fine on a small list but throws a
# real "field is not indexed" error from SharePoint once the list passes
# the ~5,000-item list view threshold.
Set-PnPField -List "WorkHub_TimeOff" -Identity "EmployeeId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
# EmployeeName is written alongside EmployeeId on every new request — the
# Team leave timeline and Pending Approvals list group/display by name.
Add-PnPField -List "WorkHub_TimeOff" -DisplayName "EmployeeName" -InternalName "EmployeeName" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TimeOff" -DisplayName "Type" -InternalName "Type" -Type Choice -Choices "Vacation","Sick leave","Personal","Holiday" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TimeOff" -DisplayName "StartDate" -InternalName "StartDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TimeOff" -DisplayName "EndDate" -InternalName "EndDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TimeOff" -DisplayName "Days" -InternalName "Days" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TimeOff" -DisplayName "Notes" -InternalName "Notes" -Type Note -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TimeOff" -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Pending","Approved","Rejected" -AddToDefaultView -ErrorAction SilentlyContinue

Write-Host "Creating WorkHub_TimeOffBalances..."
New-PnPList -Title "WorkHub_TimeOffBalances" -Template GenericList -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TimeOffBalances" -DisplayName "EmployeeId" -InternalName "EmployeeId" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
Set-PnPField -List "WorkHub_TimeOffBalances" -Identity "EmployeeId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TimeOffBalances" -DisplayName "AccruedDays" -InternalName "AccruedDays" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TimeOffBalances" -DisplayName "UsedDays" -InternalName "UsedDays" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TimeOffBalances" -DisplayName "AvailableDays" -InternalName "AvailableDays" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue

Write-Host "Creating WorkHub_TeamMembers..."
New-PnPList -Title "WorkHub_TeamMembers" -Template GenericList -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TeamMembers" -DisplayName "Email" -InternalName "Email" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TeamMembers" -DisplayName "Role" -InternalName "Role" -Type Choice -Choices "Owner","Admin","Member" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TeamMembers" -DisplayName "Group" -InternalName "Group" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List "WorkHub_TeamMembers" -DisplayName "BillableRate" -InternalName "BillableRate" -Type Currency -AddToDefaultView -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "Done. Verify columns in the SharePoint UI, then grant the WorkHub-Members / WorkHub-Admins groups access to this site."
Write-Host "Next: run .\seed-sharepoint-data.ps1 to populate starter Clients, Projects, and Team Members."