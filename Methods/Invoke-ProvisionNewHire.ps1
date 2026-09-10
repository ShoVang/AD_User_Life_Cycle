function Invoke-ProvisionNewHire {
    param(
        [Parameter(Mandatory)]
        $PendingItem,

        [Parameter(Mandatory)]
        [array]$Rows
    )

    $Row = $PendingItem.Row
    $Index = $PendingItem.Index
    $RowSource = if ($PendingItem.Source) { $PendingItem.Source } else { 'Active' }
    $DisplayName = "$(Get-RowField -Row $Row -Names @('FirstName', 'First Name')) $(Get-RowField -Row $Row -Names @('LastName', 'Last Name'))"
    $RemovedFromActive = $RowSource -eq 'Processed'

    Write-Log "--- Processing $DisplayName (source: $RowSource) ---"

    try {
        if ($RemovedFromActive) {
            $sam = Get-RowField -Row $Row -Names @('Username')
            if ([string]::IsNullOrWhiteSpace($sam)) {
                throw "Staged row on Processed tab is missing Username for $DisplayName"
            }

            Write-Log "Resuming staged hire $sam from '$ProcessedWorksheetName' tab"
            $User = Get-ADUser -Identity $sam -Properties Department, EmployeeID, Description, DistinguishedName, SamAccountName
        } else {
            $StageResult = Invoke-StageUser -Row $Row

            if ($StageResult.Status -eq 'Skip') {
                Mark-HireRowSkipped -Rows $Rows -Index $Index -SamAccountName $StageResult.SamAccountName
                return $null
            }

            $User = $StageResult
            if (-not $User.DistinguishedName) {
                $User = Get-ADUser -Identity $User.SamAccountName -Properties Department, EmployeeID, Description, DistinguishedName, SamAccountName
            }

            Set-RowProperty -Row $Row -Name 'Processed' -Value 'Staged'
            Set-RowProperty -Row $Row -Name 'Username' -Value $User.SamAccountName
            Set-RowProperty -Row $Row -Name 'StagedDate' -Value (Get-Date -Format 'yyyy-MM-dd HH:mm')
            Move-HireRowToProcessedSheet -Rows $Rows -Index $Index
            $RemovedFromActive = $true
        }

        $Mapping = Resolve-DepartmentMapping -Department $User.Department
        if (-not $Mapping) {
            Write-Log "$($User.SamAccountName): no mapping for department '$($User.Department)' - staying in 1NewUserStaging" "WARN"
            Set-RowProperty -Row $Row -Name 'ErrorMessage' -Value "No mapping for department '$($User.Department)'"
            Update-HireRowOnProcessedSheet -Row $Row
            Invoke-ReturnToStaging -User $User
            return $null
        }

        $EmployeeID = Invoke-AssignEmployeeId -User $User
        $SortResult = Invoke-SortUser -User $User -EmployeeID $EmployeeID

        if (-not $SortResult.Success) {
            Write-Log "$($User.SamAccountName): $($SortResult.Message) - returning to 1NewUserStaging" "WARN"
            Set-RowProperty -Row $Row -Name 'ErrorMessage' -Value $SortResult.Message
            Update-HireRowOnProcessedSheet -Row $Row
            Invoke-ReturnToStaging -User $User -ClearEmployeeId
            return $null
        }

        Set-RowProperty -Row $Row -Name 'Processed' -Value 'Processed'
        Set-RowProperty -Row $Row -Name 'EmployeeID' -Value $EmployeeID
        Set-RowProperty -Row $Row -Name 'ProcessedDate' -Value (Get-Date -Format 'yyyy-MM-dd HH:mm')
        Set-RowProperty -Row $Row -Name 'ErrorMessage' -Value ''
        Update-HireRowOnProcessedSheet -Row $Row

        return "PROCESSED: $DisplayName - EmployeeID $EmployeeID - Dept $($SortResult.Department) - Groups: $($SortResult.Groups -join ', ')"

    } catch {
        Write-Log "FAILED processing $DisplayName : $_" "ERROR"

        if ($_.Exception.Message -match 'already in use|already exists|83005|UNIQUE') {
            $sam = Get-RowField -Row $Row -Names @('Username')
            if ([string]::IsNullOrWhiteSpace($sam)) {
                $sam = New-Username -First (Get-RowField -Row $Row -Names @('FirstName', 'First Name')) `
                                    -Last (Get-RowField -Row $Row -Names @('LastName', 'Last Name'))
            }

            if ($RemovedFromActive) {
                Set-RowProperty -Row $Row -Name 'Processed' -Value 'Skipped'
                Set-RowProperty -Row $Row -Name 'Username' -Value $sam
                Set-RowProperty -Row $Row -Name 'SkipReason' -Value $_.Exception.Message
                Update-HireRowOnProcessedSheet -Row $Row
                Write-Log "Marked $sam as Skipped on '$ProcessedWorksheetName' tab"
            } else {
                Mark-HireRowSkipped -Rows $Rows -Index $Index -SamAccountName $sam -Reason $_.Exception.Message
            }
            return $null
        }

        Set-RowProperty -Row $Row -Name 'Processed' -Value 'Failed'
        Set-RowProperty -Row $Row -Name 'FailedDate' -Value (Get-Date -Format 'yyyy-MM-dd HH:mm')
        Set-RowProperty -Row $Row -Name 'ErrorMessage' -Value $_.Exception.Message

        if ($RemovedFromActive) {
            Update-HireRowOnProcessedSheet -Row $Row
        } else {
            Save-SpreadsheetRows -Rows $Rows
        }
        return $null
    }
}
