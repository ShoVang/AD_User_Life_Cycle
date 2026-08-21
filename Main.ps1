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
    $lastPendingKey = $null
    $samePendingCount = 0

    while ($true) {
        $Rows = Import-SpreadsheetRows
        $PendingItems = Get-PendingNewHires -Rows $Rows

        if ($PendingItems.Count -eq 0) {
            if ($SummaryLines.Count -eq 0) {
                Write-Log "No pending new hires found"
            }
            break
        }

        $PendingItem = $PendingItems[0]
        $pendingKey = "$($PendingItem.Index)|$(Get-RowField -Row $PendingItem.Row -Names @('FirstName', 'First Name'))|$(Get-RowField -Row $PendingItem.Row -Names @('LastName', 'Last Name'))"

        if ($pendingKey -eq $lastPendingKey) {
            $samePendingCount++
            if ($samePendingCount -ge 3) {
                Write-Log "Stopping: same hire stuck in queue ($pendingKey). Check Processed column in spreadsheet." "ERROR"
                break
            }
        } else {
            $samePendingCount = 0
        }

        $lastPendingKey = $pendingKey

        $Result = Invoke-ProvisionNewHire -PendingItem $PendingItem -Rows $Rows
        if ($Result) {
            $SummaryLines += $Result
        }
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