function Connect-SharePointIfNeeded {
    if ($Script:SharePointConnected) { return }

    Import-Module PnP.PowerShell -ErrorAction Stop
    Connect-PnPOnline -Url $SharePointSiteUrl @SharePointConnectParams
    $Script:SharePointConnected = $true
    Write-Log "Connected to SharePoint: $SharePointSiteUrl"
}

function Import-SpreadsheetFromSource {
    $workbookDir = Split-Path $Script:WorkbookPath -Parent
    if (-not (Test-Path $workbookDir)) {
        New-Item -ItemType Directory -Path $workbookDir -Force | Out-Null
    }

    switch ($SpreadsheetMode) {
        'Local' {
            if ([string]::IsNullOrWhiteSpace($SpreadsheetPath)) {
                throw "SpreadsheetMode is Local but `$SpreadsheetPath is not set."
            }
            if (-not (Test-Path $SpreadsheetPath)) {
                throw "Spreadsheet not found at $SpreadsheetPath"
            }

            if ($SpreadsheetPath -ne $Script:WorkbookPath) {
                Copy-Item -Path $SpreadsheetPath -Destination $Script:WorkbookPath -Force
            }

            Write-Log "Loaded spreadsheet from local path: $SpreadsheetPath"
        }

        'Url' {
            if ([string]::IsNullOrWhiteSpace($SpreadsheetUrl)) {
                throw "SpreadsheetMode is Url but `$SpreadsheetUrl is not set."
            }

            Invoke-WebRequest -Uri $SpreadsheetUrl -OutFile $Script:WorkbookPath -UseBasicParsing
            Write-Log "Downloaded spreadsheet from URL"
        }

        'SharePoint' {
            if ([string]::IsNullOrWhiteSpace($SharePointSiteUrl) -or [string]::IsNullOrWhiteSpace($SharePointServerRelativeUrl)) {
                throw "SpreadsheetMode is SharePoint but site URL or file path is not set."
            }

            Connect-SharePointIfNeeded
            $fileName = Split-Path $Script:WorkbookPath -Leaf
            Get-PnPFile -Url $SharePointServerRelativeUrl -Path $workbookDir -Filename $fileName -AsFile -Force
            Write-Log "Downloaded spreadsheet from SharePoint: $SharePointServerRelativeUrl"
        }

        default {
            throw "Unknown SpreadsheetMode '$SpreadsheetMode'. Use Local, Url, or SharePoint."
        }
    }
}

function Publish-SpreadsheetToSource {
    if (-not (Test-Path $Script:WorkbookPath)) {
        Write-Log "Workbook not found for upload: $Script:WorkbookPath" "ERROR"
        return
    }

    switch ($SpreadsheetMode) {
        'Local' {
            if (-not [string]::IsNullOrWhiteSpace($SpreadsheetPath) -and $SpreadsheetPath -ne $Script:WorkbookPath) {
                Copy-Item -Path $Script:WorkbookPath -Destination $SpreadsheetPath -Force
                Write-Log "Saved spreadsheet back to local path: $SpreadsheetPath"
            }
        }

        'Url' {
            if ([string]::IsNullOrWhiteSpace($SpreadsheetUploadUrl)) {
                Write-Log "SpreadsheetUploadUrl is not set; workbook updated locally only" "WARN"
                return
            }

            $bytes = [System.IO.File]::ReadAllBytes($Script:WorkbookPath)
            Invoke-WebRequest -Uri $SpreadsheetUploadUrl -Method Put -Body $bytes -UseBasicParsing | Out-Null
            Write-Log "Uploaded spreadsheet to URL"
        }

        'SharePoint' {
            Connect-SharePointIfNeeded

            $folderUrl = Split-Path $SharePointServerRelativeUrl -Parent
            $fileName  = Split-Path $SharePointServerRelativeUrl -Leaf
            Add-PnPFile -Path $Script:WorkbookPath -Folder $folderUrl -NewFileName $fileName
            Write-Log "Uploaded spreadsheet to SharePoint: $SharePointServerRelativeUrl"
        }
    }
}

function Get-WorkbookWorksheetNames {
    param([string]$Path)

    $package = Open-ExcelPackage -Path $Path
    try {
        return @($package.Workbook.Worksheets | ForEach-Object { $_.Name })
    } finally {
        Close-ExcelPackage $package
    }
}

function Import-WorksheetRows {
    param([Parameter(Mandatory)][string]$SheetName)

    $availableSheets = Get-WorkbookWorksheetNames -Path $Script:WorkbookPath
    if ($availableSheets -notcontains $SheetName) {
        return @()
    }

    $rows = Import-Excel -Path $Script:WorkbookPath -WorksheetName $SheetName -StartRow $SpreadsheetStartRow
    return @($rows)
}

function Import-SpreadsheetRows {
    $availableSheets = Get-WorkbookWorksheetNames -Path $Script:WorkbookPath
    if ($availableSheets -notcontains $WorksheetName) {
        throw "Worksheet '$WorksheetName' not found. Available tabs: $($availableSheets -join ', ')"
    }

    return Import-WorksheetRows -SheetName $WorksheetName
}

function Initialize-SpreadsheetRowColumns {
    param([array]$Rows)

    $trackingColumns = @(
        'Processed', 'Username', 'EmployeeID', 'StagedDate', 'ProcessedDate',
        'FailedDate', 'ErrorMessage', 'SkipReason', 'Status'
    )

    foreach ($row in $Rows) {
        foreach ($column in $trackingColumns) {
            if ($row.PSObject.Properties.Name -notcontains $column) {
                Set-RowProperty -Row $row -Name $column -Value ''
            }
        }
    }

    return $Rows
}

function Clear-WorksheetExtraDataRows {
    param(
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][int]$DataRowCount
    )

    $package = Open-ExcelPackage -Path $Script:WorkbookPath
    try {
        $ws = $package.Workbook.Worksheets[$SheetName]
        if (-not $ws -or $null -eq $ws.Dimension) { return }

        $lastWrittenRow = $SpreadsheetStartRow + $DataRowCount
        $usedEndRow = $ws.Dimension.End.Row
        $usedEndCol = $ws.Dimension.End.Column

        if ($usedEndRow -gt $lastWrittenRow) {
            $ws.Cells[$lastWrittenRow + 1, 1, $usedEndRow, $usedEndCol].Clear()
        }

        Close-ExcelPackage $package -Save $true
    } catch {
        Close-ExcelPackage $package
        throw
    }
}

function Save-WorksheetRows {
    param(
        [array]$Rows,
        [Parameter(Mandatory)][string]$SheetName
    )

    Initialize-SpreadsheetRowColumns -Rows $Rows | Out-Null
    $Rows | Export-Excel -Path $Script:WorkbookPath -WorksheetName $SheetName `
        -StartRow $SpreadsheetStartRow -AutoSize
    Clear-WorksheetExtraDataRows -SheetName $SheetName -DataRowCount @($Rows).Count
}

function Move-HireRowToProcessedSheet {
    param(
        [Parameter(Mandatory)][array]$Rows,
        [Parameter(Mandatory)][int]$Index
    )

    $rowToMove = $Rows[$Index]
    $firstName = Get-RowField -Row $rowToMove -Names @('FirstName', 'First Name')
    $lastName  = Get-RowField -Row $rowToMove -Names @('LastName', 'Last Name')

    $processedRows = Import-WorksheetRows -SheetName $ProcessedWorksheetName
    $processedRows = @($processedRows) + @($rowToMove)

    $activeRows = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $Rows.Count; $i++) {
        if ($i -ne $Index) {
            $activeRows.Add($Rows[$i])
        }
    }

    Save-WorksheetRows -Rows $processedRows -SheetName $ProcessedWorksheetName
    Save-WorksheetRows -Rows @($activeRows) -SheetName $WorksheetName
    Publish-SpreadsheetToSource

    Write-Log "Moved $firstName $lastName from Active to '$ProcessedWorksheetName' tab (removed from Active)"
}

function Move-CompletedActiveRowsToProcessedSheet {
    $rows = Import-SpreadsheetRows
    if ($rows.Count -eq 0) { return }

    $toMove = [System.Collections.Generic.List[object]]::new()
    $toKeep = [System.Collections.Generic.List[object]]::new()

    foreach ($row in $rows) {
        $processed = Get-RowField -Row $row -Names @('Processed')
        if ($processed -ieq 'Processed') {
            $toMove.Add($row)
        } else {
            $toKeep.Add($row)
        }
    }

    if ($toMove.Count -eq 0) { return }

    $processedRows = Import-WorksheetRows -SheetName $ProcessedWorksheetName
    foreach ($row in $toMove) {
        $processedRows += $row
    }

    Save-WorksheetRows -Rows @($processedRows) -SheetName $ProcessedWorksheetName
    Save-WorksheetRows -Rows @($toKeep) -SheetName $WorksheetName
    Publish-SpreadsheetToSource

    Write-Log "Moved $($toMove.Count) completed row(s) from Active to '$ProcessedWorksheetName' (removed from Active)"
}

function Write-SpreadsheetDiagnostics {
    $availableSheets = Get-WorkbookWorksheetNames -Path $Script:WorkbookPath
    $rows = Import-SpreadsheetRows

    Write-Log "Spreadsheet diagnostics: path=$Script:WorkbookPath, worksheet='$WorksheetName', startRow=$SpreadsheetStartRow, importedRows=$($rows.Count)"
    Write-Log "Available worksheet tabs: $($availableSheets -join ', ')"

    if ($rows.Count -eq 0) {
        Write-Log "No data rows were imported. Check `$WorksheetName and `$SpreadsheetStartRow." "WARN"
        return
    }

    $columns = ($rows[0].PSObject.Properties | ForEach-Object { $_.Name }) -join ', '
    Write-Log "Detected columns: $columns"

    $sample = $rows[0]
    Write-Log ("Sample row 1: FirstName='{0}', LastName='{1}', Department='{2}', Processed='{3}', Status='{4}'" -f `
        (Get-RowField -Row $sample -Names @('FirstName', 'First Name')), `
        (Get-RowField -Row $sample -Names @('LastName', 'Last Name')), `
        (Get-RowField -Row $sample -Names @('Department', 'Departmer', 'Dept')), `
        (Get-RowField -Row $sample -Names @('Processed')), `
        (Get-RowField -Row $sample -Names @('Status')))
}

function Save-SpreadsheetRows {
    param([array]$Rows)

    Save-WorksheetRows -Rows $Rows -SheetName $WorksheetName
    Publish-SpreadsheetToSource
}
