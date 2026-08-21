function Invoke-StageUser {
    param(
        [Parameter(Mandatory)]
        $Row
    )

    $firstName = Get-RowField -Row $Row -Names @('FirstName', 'First Name')
    $lastName  = Get-RowField -Row $Row -Names @('LastName', 'Last Name')
    $username  = Get-RowField -Row $Row -Names @('Username')

    if ([string]::IsNullOrWhiteSpace($username)) {
        $username = New-Username -First $firstName -Last $lastName
    }

    $Existing = Get-ADUser -Filter "SamAccountName -eq '$username'" `
        -Properties Department, EmployeeID, Description, DistinguishedName -ErrorAction SilentlyContinue

    if ($Existing) {
        Write-Log "Account $username already exists in AD - skipping row"
        return [PSCustomObject]@{
            Skip           = $true
            SamAccountName = $username
        }
    }

    $UPN      = "$username@$DefaultDomain"
    $PlainPW  = New-RandomPassword -Length $DefaultPassLen
    $SecurePW = ConvertTo-SecureString $PlainPW -AsPlainText -Force
    $title      = Get-RowField -Row $Row -Names @('Title', 'JobTitle', 'Job Title')
    $department = Get-RowField -Row $Row -Names @('Department', 'Departmer', 'Dept')

    New-ADUser -Name "$firstName $lastName" `
        -GivenName $firstName `
        -Surname $lastName `
        -SamAccountName $username `
        -UserPrincipalName $UPN `
        -Description $title `
        -Department $department `
        -Path $StagingOU `
        -AccountPassword $SecurePW `
        -Enabled $true `
        -ChangePasswordAtLogon $true

    Write-Log "Staged $username ($firstName $lastName) in 1NewUserStaging - Dept: $department"

    return [PSCustomObject]@{
        Skip              = $false
        SamAccountName    = $username
        Username          = $username
        Name              = "$firstName $lastName"
        Department        = $department
        DistinguishedName = (Get-ADUser -Identity $username).DistinguishedName
        EmployeeID        = $null
    }
}
