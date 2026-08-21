<#
.SYNOPSIS
    Entry point. Dot-sources Functions.ps1 (which loads Private/*.ps1) and runs
    the provisioning loop one user at a time until no pending hires remain.

.NOTES
    Run as a scheduled task under a dedicated service account
    (e.g. svc-adprovisioning) with delegated rights on the staging OU and
    all target department OUs.
#>

. "$PSScriptRoot\Functions.ps1"

function Main {
    Write-Log "=== Provisioning run started (SpreadsheetMode: $SpreadsheetMode) ==="

    try {
        Import-SpreadsheetFromSource
    } catch {
        Write-Log "Failed to load spreadsheet: $_" "ERROR"
        Write-Log "=== Provisioning run finished ==="
        return
    }

    $SummaryLines = @()
    $Rows = Import-SpreadsheetRows
    $PendingItems = Get-PendingNewHires

    if ($PendingItems.Count -eq 0) {
        Write-Log "No pending new hires found"
        Write-Log "=== Provisioning run finished ==="
        return
    }

    # Process one user at a time. Re-reading the pending list each pass
    # (rather than looping over a snapshot) means the run always reflects
    # the latest state of the spreadsheet after each write-back.
    while ($PendingItems.Count -gt 0) {
        $PendingItem = $PendingItems[0]

        $Result = Invoke-ProvisionNewHire -PendingItem $PendingItem -Rows $Rows
        if ($Result) {
            $SummaryLines += $Result
        }

        $Rows = Import-SpreadsheetRows
        $PendingItems = Get-PendingNewHires
    }

    if ($SummaryLines.Count -gt 0) {
        $Body = $SummaryLines -join "`n"
        Send-MailMessage -SmtpServer $SmtpServer -From $MailFrom -To $MailTo `
            -Subject "AD Provisioning Run - $(Get-Date -Format 'yyyy-MM-dd')" -Body $Body
        Write-Log "Summary email sent to $MailTo"
    }

    Write-Log "=== Provisioning run finished ==="
}

Main