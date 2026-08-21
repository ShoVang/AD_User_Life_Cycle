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
        $StageResult = Invoke-StageUser -Row $Row

        if ($StageResult.Status -eq 'Skip') {
            Mark-HireRowSkipped -Rows $Rows -Index $Index -SamAccountName $StageResult.SamAccountName
            return $null
        }

        $User = $StageResult
        if (-not $User.DistinguishedName) {
            $User = Get-ADUser -Identity $User.SamAccountName -Properties Department, EmployeeID, Description
        }

        Set-RowProperty -Row $Rows[$Index] -Name 'Processed' -Value 'Staged'
        Set-RowProperty -Row $Rows[$Index] -Name 'Username' -Value $User.SamAccountName
        Set-RowProperty -Row $Rows[$Index] -Name 'StagedDate' -Value (Get-Date -Format 'yyyy-MM-dd HH:mm')
        Save-SpreadsheetRows -Rows $Rows

        $Mapping = Resolve-DepartmentMapping -Department $User.Department
        if (-not $Mapping) {
            Write-Log "$($User.SamAccountName): no mapping for department '$($User.Department)' - staying in 1NewUserStaging" "WARN"
            Invoke-ReturnToStaging -User $User
            return $null
        }

        $EmployeeID = Invoke-AssignEmployeeId -User $User
        $SortResult = Invoke-SortUser -User $User -EmployeeID $EmployeeID

        if (-not $SortResult.Success) {
            Write-Log "$($User.SamAccountName): $($SortResult.Message) - returning to 1NewUserStaging" "WARN"
            Invoke-ReturnToStaging -User $User -ClearEmployeeId
            return $null
        }

        Set-RowProperty -Row $Rows[$Index] -Name 'Processed' -Value 'Processed'
        Set-RowProperty -Row $Rows[$Index] -Name 'EmployeeID' -Value $EmployeeID
        Set-RowProperty -Row $Rows[$Index] -Name 'ProcessedDate' -Value (Get-Date -Format 'yyyy-MM-dd HH:mm')
        Save-SpreadsheetRows -Rows $Rows

        return "PROCESSED: $DisplayName - EmployeeID $EmployeeID - Dept $($SortResult.Department) - Groups: $($SortResult.Groups -join ', ')"

    } catch {
        Write-Log "FAILED processing $DisplayName : $_" "ERROR"

        if ($_.Exception.Message -match 'already in use|already exists|83005|UNIQUE') {
            $sam = Get-RowField -Row $Row -Names @('Username')
            if ([string]::IsNullOrWhiteSpace($sam)) {
                $sam = New-Username -First (Get-RowField -Row $Row -Names @('FirstName', 'First Name')) `
                                    -Last (Get-RowField -Row $Row -Names @('LastName', 'Last Name'))
            }
            Mark-HireRowSkipped -Rows $Rows -Index $Index -SamAccountName $sam -Reason $_.Exception.Message
            return $null
        }

        Set-RowProperty -Row $Rows[$Index] -Name 'Processed' -Value 'Failed'
        Set-RowProperty -Row $Rows[$Index] -Name 'FailedDate' -Value (Get-Date -Format 'yyyy-MM-dd HH:mm')
        Set-RowProperty -Row $Rows[$Index] -Name 'ErrorMessage' -Value $_.Exception.Message
        Save-SpreadsheetRows -Rows $Rows
        return $null
    }
}
