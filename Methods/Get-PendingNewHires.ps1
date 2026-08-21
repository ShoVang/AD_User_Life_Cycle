function Get-PendingNewHires {
    param([array]$Rows)

    if (-not (Test-Path $Script:WorkbookPath)) {
        Write-Log "Spreadsheet not found at $Script:WorkbookPath" "ERROR"
        return @()
    }

    if (-not $Rows) {
        $Rows = Import-SpreadsheetRows
    }
    $Pending = [System.Collections.Generic.List[object]]::new()

    for ($i = 0; $i -lt $Rows.Count; $i++) {
        $Row = $Rows[$i]
        if (Test-RowShouldSkip -Row $Row) { continue }
        if (Test-RowIsExample -Row $Row) { continue }

        $firstName = Get-RowField -Row $Row -Names @('FirstName', 'First Name')
        $lastName  = Get-RowField -Row $Row -Names @('LastName', 'Last Name')
        if ([string]::IsNullOrWhiteSpace($firstName) -or [string]::IsNullOrWhiteSpace($lastName)) { continue }

        $Pending.Add([PSCustomObject]@{
            Index = $i
            Row   = $Row
        })
    }

    if ($Pending.Count -eq 0) {
        Write-SpreadsheetDiagnostics
    }

    return $Pending
}
