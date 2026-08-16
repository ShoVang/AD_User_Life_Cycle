<# 0111 1
.SYNOPSIS
    AD provisioning functions and config. Dot-sourced by Main.ps1.

.DESCRIPTION
    Defines config, helpers, and provisioning jobs. Does not run on its own.
    Run Main.ps1 to execute the provisioning loop.

.NOTES
    Run as a scheduled task under a dedicated service account
    (e.g. svc-adprovisioning) with delegated rights on the staging OU and all
    target department OUs.

    Spreadsheet source modes:
        Local      - file on disk or UNC (\\server\share\file.xlsx)
        Url        - direct HTTPS download link (must return the .xlsx bytes)
        SharePoint - SharePoint Online site + file path (recommended for prod)

    Install SharePoint support:  Install-Module PnP.PowerShell
#>

# ============================================================
# CONFIG - edit these for your environment
# ============================================================

# Pick ONE source mode: Local | Url | SharePoint
$SpreadsheetMode   = "Local"

# Local mode - file path on this machine or UNC share
$SpreadsheetPath    = "C:\HR_Drive\HR_NewHires.xlsx"

# Url mode - direct HTTPS link that downloads the .xlsx file
# SharePoint "copy link" pages won't work; you need a direct download URL
# or use SharePoint mode instead.
$SpreadsheetUrl     = "https://yourtenant.sharepoint.com/.../HR_NewHires.xlsx"
$SpreadsheetUploadUrl = ""   # optional PUT/upload URL; leave blank if read-only

# SharePoint mode - recommended for production
$SharePointSiteUrl          = "https://yourtenant.sharepoint.com/sites/HR"
$SharePointServerRelativeUrl = "/sites/HR/Shared Documents/HR_NewHires.xlsx"

# How the script signs in to SharePoint (pick one approach):
#   Testing:     @{ Interactive = $true }
#   Scheduled:   @{ ClientId = "app-id"; Thumbprint = "cert-thumbprint"; Tenant = "yourtenant.onmicrosoft.com" }
$SharePointConnectParams = @{ Interactive = $true }

# Local working copy - always used for Import-Excel / Export-Excel
$LocalWorkbookPath  = "C:\ProvisioningLogs\HR_NewHires.xlsx"
$WorksheetName      = "Active"
$SpreadsheetStartRow = 4   # row where column headers live (rows 1-3 are template title/instructions)
$LogPath           = "C:\ProvisioningLogs\provisioning_$(Get-Date -Format 'yyyyMMdd').log"
$DefaultDomain     = "mydomain.com"
$DefaultPassLen    = 16

$StagingOU         = "OU=1NewUserStaging,OU=Users,DC=mydomain,DC=com"
$EmployeeIDFile    = "C:\ProvisioningLogs\next_employee_id.txt"   # tracks the next sequential ID
$EmployeeIDPrefix  = "EMP"
$EmployeeIDSeed    = 1001

$SmtpServer        = "smtp.yourdomain.com"
$MailFrom          = "ad-automation@yourdomain.com"
$MailTo            = "it-team@yourdomain.com"

# Department -> OU + Groups mapping. EDIT THIS to match your real AD structure.
$DeptMap = @{
    "Sales"      = @{ OU = "OU=Sales,OU=Users,DC=mydomain,DC=com";      Groups = @("Sales-Team","VPN-Users","CRM-Access") }
    "Accounting" = @{ OU = "OU=Finance,OU=Users,DC=mydomain,DC=com";    Groups = @("Finance-Team","ERP-Access") }
    "IT"         = @{ OU = "OU=IT,OU=Users,DC=mydomain,DC=com";         Groups = @("IT-Team","VPN-Users") }
}

# ============================================================
# SETUP
# ============================================================

Import-Module ActiveDirectory
Import-Module ImportExcel -ErrorAction Stop   # Install-Module ImportExcel if missing

$Script:WorkbookPath = $LocalWorkbookPath
$Script:SharePointConnected = $false

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
    Add-Content -Path $LogPath -Value $line
    Write-Host $line
}

function New-RandomPassword {
    param([int]$Length = 16)
    $chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%'
    -join (1..$Length | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
}

function New-Username {
    param([string]$First, [string]$Last)
    $base = ("$($First.Substring(0,1))$Last" -replace '[^a-zA-Z]', '').ToLower()
    $candidate = $base
    $i = 1
    while (Get-ADUser -Filter "SamAccountName -eq '$candidate'" -ErrorAction SilentlyContinue) {
        $i++
        $candidate = "$base$i"
    }
    return $candidate
}

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

function Get-RowField {
    param(
        $Row,
        [Parameter(Mandatory)]
        [string[]]$Names
    )

    foreach ($name in $Names) {
        $prop = $Row.PSObject.Properties | Where-Object { $_.Name -ieq $name } | Select-Object -First 1
        if ($prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
            return [string]$prop.Value.Trim()
        }
    }

    return $null
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

function Test-RowAlreadyProcessed {
    param($Row)

    if ($Row.Processed -eq "Processed") { return $true }
    if ($Row.Status -eq "Processed") { return $true }

    return $false
}

function Test-RowIsExample {
    param($Row)

    $firstName = Get-RowField -Row $Row -Names @('FirstName', 'First Name')
    $lastName  = Get-RowField -Row $Row -Names @('LastName', 'Last Name')
    $notes     = Get-RowField -Row $Row -Names @('Notes', 'Note')

    if ($firstName -eq "First" -and $lastName -eq "Last") { return $true }
    if ($notes -match 'example row') { return $true }

    return $false
}

# ============================================================
# JOB: Grab user
# ============================================================

function Get-PendingNewHires {
    if (-not (Test-Path $Script:WorkbookPath)) {
        Write-Log "Spreadsheet not found at $Script:WorkbookPath" "ERROR"
        return @()
    }

    $Rows = Import-SpreadsheetRows
    $Pending = [System.Collections.Generic.List[object]]::new()

    for ($i = 0; $i -lt $Rows.Count; $i++) {
        $Row = $Rows[$i]
        if (Test-RowAlreadyProcessed -Row $Row) { continue }
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

# ============================================================
# JOB: Stage user -> 1NewUserStaging
# ============================================================

function Invoke-StageUser {
    param(
        [Parameter(Mandatory)]
        $Row
    )

    if ($Row.Processed -eq "Staged" -and -not [string]::IsNullOrWhiteSpace($Row.Username)) {
        $Existing = Get-ADUser -Filter "SamAccountName -eq '$($Row.Username)'" `
            -Properties Department, EmployeeID, Description -ErrorAction SilentlyContinue

        if ($Existing) {
            Write-Log "Located existing staged account $($Row.Username)"
            return $Existing
        }

        Write-Log "Spreadsheet says Staged but AD account '$($Row.Username)' not found" "WARN"
    }

    $Username = New-Username -First (Get-RowField -Row $Row -Names @('FirstName', 'First Name')) `
                             -Last (Get-RowField -Row $Row -Names @('LastName', 'Last Name'))
    $UPN      = "$Username@$DefaultDomain"
    $PlainPW  = New-RandomPassword -Length $DefaultPassLen
    $SecurePW = ConvertTo-SecureString $PlainPW -AsPlainText -Force
    $firstName = Get-RowField -Row $Row -Names @('FirstName', 'First Name')
    $lastName  = Get-RowField -Row $Row -Names @('LastName', 'Last Name')
    $title     = Get-RowField -Row $Row -Names @('Title', 'JobTitle', 'Job Title')
    $department = Get-RowField -Row $Row -Names @('Department', 'Departmer', 'Dept')

    New-ADUser -Name "$firstName $lastName" `
        -GivenName $firstName `
        -Surname $lastName `
        -SamAccountName $Username `
        -UserPrincipalName $UPN `
        -Description $title `
        -Department $department `
        -Path $StagingOU `
        -AccountPassword $SecurePW `
        -Enabled $true `
        -ChangePasswordAtLogon $true

    Write-Log "Staged $Username ($firstName $lastName) in 1NewUserStaging - Dept: $department"

    return [PSCustomObject]@{
        SamAccountName      = $Username
        Username            = $Username
        Name                = "$firstName $lastName"
        Department          = $department
        DistinguishedName   = (Get-ADUser -Identity $Username).DistinguishedName
        EmployeeID          = $null
    }
}

# ============================================================
# JOB: Assign Employee ID
# ============================================================

function Get-NextEmployeeID {
    if (-not (Test-Path $EmployeeIDFile)) {
        Set-Content -Path $EmployeeIDFile -Value $EmployeeIDSeed
    }
    $current = [int](Get-Content -Path $EmployeeIDFile -Raw)
    $nextValue = $current + 1
    Set-Content -Path $EmployeeIDFile -Value $nextValue
    return "$EmployeeIDPrefix$current"
}

function Invoke-AssignEmployeeId {
    param(
        [Parameter(Mandatory)]
        $User
    )

    if (-not [string]::IsNullOrWhiteSpace($User.EmployeeID)) {
        Write-Log "$($User.SamAccountName) already has EmployeeID $($User.EmployeeID)"
        return $User.EmployeeID
    }

    $EmployeeID = Get-NextEmployeeID
    Set-ADUser -Identity $User.SamAccountName -EmployeeID $EmployeeID
    Write-Log "Assigned EmployeeID $EmployeeID to $($User.SamAccountName)"
    return $EmployeeID
}

function Invoke-ClearEmployeeId {
    param(
        [Parameter(Mandatory)]
        [string]$SamAccountName
    )

    Set-ADUser -Identity $SamAccountName -Clear EmployeeID
    Write-Log "Cleared EmployeeID on $SamAccountName"
}

# ============================================================
# JOB: Sort user by department / organization
# ============================================================

function Resolve-DepartmentMapping {
    param([string]$Department)

    if ([string]::IsNullOrWhiteSpace($Department)) {
        return $null
    }

    if ($DeptMap.ContainsKey($Department)) {
        return $DeptMap[$Department]
    }

    return $null
}

function Invoke-SortUser {
    param(
        [Parameter(Mandatory)]
        $User,

        [Parameter(Mandatory)]
        [string]$EmployeeID
    )

    $Dept = $User.Department
    $Mapping = Resolve-DepartmentMapping -Department $Dept

    if (-not $Mapping) {
        return [PSCustomObject]@{
            Success = $false
            Reason  = "UnknownDepartment"
            Message = "No mapping for department '$Dept'"
        }
    }

    try {
        Move-ADObject -Identity $User.DistinguishedName -TargetPath $Mapping.OU

        foreach ($Group in $Mapping.Groups) {
            try {
                Add-ADGroupMember -Identity $Group -Members $User.SamAccountName
            } catch {
                Write-Log "Failed to add $($User.SamAccountName) to group $Group : $_" "ERROR"
            }
        }

        Write-Log "Sorted $($User.SamAccountName) -> $Dept, EmployeeID $EmployeeID, OU $($Mapping.OU)"
        return [PSCustomObject]@{
            Success    = $true
            Department = $Dept
            EmployeeID = $EmployeeID
            OU         = $Mapping.OU
            Groups     = $Mapping.Groups
        }

    } catch {
        return [PSCustomObject]@{
            Success = $false
            Reason  = "SortFailed"
            Message = $_.Exception.Message
        }
    }
}

# ============================================================
# JOB: Return user to 1NewUserStaging
# ============================================================

function Invoke-ReturnToStaging {
    param(
        [Parameter(Mandatory)]
        $User,

        [switch]$ClearEmployeeId
    )

    $Current = Get-ADUser -Identity $User.SamAccountName -Properties DistinguishedName, EmployeeID

    if ($ClearEmployeeId -and -not [string]::IsNullOrWhiteSpace($Current.EmployeeID)) {
        Invoke-ClearEmployeeId -SamAccountName $User.SamAccountName
    }

    if ($Current.DistinguishedName -notlike "*$StagingOU") {
        try {
            Move-ADObject -Identity $Current.DistinguishedName -TargetPath $StagingOU
            Write-Log "Returned $($User.SamAccountName) to 1NewUserStaging"
        } catch {
            Write-Log "Failed to return $($User.SamAccountName) to staging: $_" "ERROR"
        }
    } else {
        Write-Log "$($User.SamAccountName) remains in 1NewUserStaging"
    }
}

# ============================================================
# Per-user pipeline (all jobs for one hire)
# ============================================================

function Invoke-ProvisionNewHire {
    param(
        [Parameter(Mandatory)]
        $PendingItem,

        [Parameter(Mandatory)]
        [array]$Rows
    )

    $Row = $PendingItem.Row
    $Index = $PendingItem.Index
    $DisplayName = "$(Get-RowField -Row $Row -Names @('FirstName', 'First Name')) $(Get-RowField -Row $Row -Names @('LastName', 'Last Name'))"

    Write-Log "--- Processing $DisplayName ---"

    try {
        # Job 1: Grab (caller selected this row)
        # Job 2: Stage
        $User = Invoke-StageUser -Row $Row
        if (-not $User.DistinguishedName) {
            $User = Get-ADUser -Identity $User.SamAccountName -Properties Department, EmployeeID, Description
        }

        $Rows[$Index].Processed = "Staged"
        $Rows[$Index] | Add-Member -NotePropertyName Username -NotePropertyValue $User.SamAccountName -Force
        $Rows[$Index] | Add-Member -NotePropertyName StagedDate -NotePropertyValue (Get-Date -Format 'yyyy-MM-dd HH:mm') -Force
        Save-SpreadsheetRows -Rows $Rows

        $Mapping = Resolve-DepartmentMapping -Department $User.Department
        if (-not $Mapping) {
            Write-Log "$($User.SamAccountName): no mapping for department '$($User.Department)' - staying in 1NewUserStaging" "WARN"
            Invoke-ReturnToStaging -User $User
            return $null
        }

        # Job 3: Assign Employee ID
        $EmployeeID = Invoke-AssignEmployeeId -User $User

        # Job 4: Sort
        $SortResult = Invoke-SortUser -User $User -EmployeeID $EmployeeID

        if (-not $SortResult.Success) {
            # Job 5: Return to staging when sort fails after ID was assigned
            Write-Log "$($User.SamAccountName): $($SortResult.Message) - returning to 1NewUserStaging" "WARN"
            Invoke-ReturnToStaging -User $User -ClearEmployeeId
            return $null
        }

        $Rows[$Index].Processed = "Processed"
        $Rows[$Index] | Add-Member -NotePropertyName EmployeeID -NotePropertyValue $EmployeeID -Force
        $Rows[$Index] | Add-Member -NotePropertyName ProcessedDate -NotePropertyValue (Get-Date -Format 'yyyy-MM-dd HH:mm') -Force
        Save-SpreadsheetRows -Rows $Rows

        return "PROCESSED: $DisplayName - EmployeeID $EmployeeID - Dept $($SortResult.Department) - Groups: $($SortResult.Groups -join ', ')"

    } catch {
        Write-Log "FAILED processing $DisplayName : $_" "ERROR"
        return $null
    }
}
