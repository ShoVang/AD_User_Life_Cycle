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
