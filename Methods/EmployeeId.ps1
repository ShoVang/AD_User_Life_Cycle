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
