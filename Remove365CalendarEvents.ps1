<#
.SYNOPSIS
  Removes all calendar events and meetings for a specified user via Microsoft Graph.

.DESCRIPTION
  Connects to Microsoft Graph, retrieves all events in the user's primary calendar,
  and deletes each one. Deleting a meeting sends cancellation messages to attendees.
#>

# 1. Import the Graph Calendar module
Import-Module Microsoft.Graph.Calendar

# 2. Sign in to Microsoft Graph with calendar read/write permissions
Connect-MgGraph -Scopes "Calendars.ReadWrite"  # :contentReference[oaicite:0]{index=0}

# 3. Prompt for the target user's UPN
$targetUPN = Read-Host -Prompt "Enter the User Principal Name (UPN) of the user"

# 4. Confirm before proceeding
Write-Host "WARNING: This will DELETE ALL events for: $targetUPN" -ForegroundColor Yellow
$confirmation = Read-Host -Prompt "Type 'Yes' to confirm"
if ($confirmation -ne 'Yes') {
    Write-Host "Operation cancelled."
    exit
}

# 5. Retrieve all events from the user's primary calendar
Write-Host "Retrieving events for $targetUPN..."
$events = Get-MgUserEvent -UserId $targetUPN -All  # 

if (-not $events) {
    Write-Host "No events found for user $targetUPN."
    exit
}

# 6. Iterate and delete each event
foreach ($event in $events) {
    try {
        Remove-MgUserEvent -UserId $targetUPN -EventId $event.Id -Confirm:$false  # 
        Write-Host "Deleted: '$($event.Subject)' (ID: $($event.Id))"
    }
    catch {
        Write-Warning "Failed to delete event $($event.Id): $_"
    }
}

Write-Host "Finished. Removed $($events.Count) events from $targetUPN's calendar."
