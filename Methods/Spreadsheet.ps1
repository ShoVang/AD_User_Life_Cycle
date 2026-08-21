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

function Import-SpreadsheetRows {
    $availableSheets = Get-WorkbookWorksheetNames -Path $Script:WorkbookPath
    if ($availableSheets -notcontains $WorksheetName) {
        throw "Worksheet '$WorksheetName' not found. Available tabs: $($availableSheets -join ', ')"
    }

    $rows = Import-Excel -Path $Script:WorkbookPath -WorksheetName $WorksheetName -StartRow $SpreadsheetStartRow
    return @($rows)
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
    $Rows | Export-Excel -Path $Script:WorkbookPath -WorksheetName $WorksheetName `
        -StartRow $SpreadsheetStartRow -AutoSize
    Publish-SpreadsheetToSource
}
