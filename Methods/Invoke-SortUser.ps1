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
