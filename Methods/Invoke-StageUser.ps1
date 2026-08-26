function Invoke-StageUser {
    param(
        [Parameter(Mandatory)]
        $Row
    )

    $firstName = Get-RowField -Row $Row -Names @('FirstName', 'First Name')
    $lastName  = Get-RowField -Row $Row -Names @('LastName', 'Last Name')
    
    # i can remove this VV
    $username  = Get-RowField -Row $Row -Names @('Username')

    $Existing = Get-ExistingAdUserForHire -FirstName $firstName -LastName $lastName -SamAccountName $username
    if ($Existing) {
        Write-Log "Account $($Existing.SamAccountName) already exists for $firstName $lastName - skipping row"
        return [PSCustomObject]@{
            Status         = 'Skip'
            SamAccountName = $Existing.SamAccountName
        }
    }

    if ([string]::IsNullOrWhiteSpace($username)) {
        $username = New-Username -First $firstName -Last $lastName
    }

    $displayName = Set-DisplayName -First $firstName -Last $lastName
    $license     = Get-RowField -Row $Row -Names @('License')
    $email       = Set-Email -First $firstName -Last $lastName -License $license
    $UPN         = if ($email) { $email } else { "$username@$DefaultDomain" }

    $PlainPW  = New-RandomPassword -Length $DefaultPassLen
    $SecurePW = ConvertTo-SecureString $PlainPW -AsPlainText -Force
    $title      = Get-RowField -Row $Row -Names @('Title', 'JobTitle', 'Job Title')
    $department = Get-RowField -Row $Row -Names @('Department', 'Departmer', 'Dept')

    $newUserParams = @{
        Name                  = $displayName
        GivenName             = $firstName
        Surname               = $lastName
        SamAccountName        = $username
        UserPrincipalName     = $UPN
        Description           = $title
        Department            = $department
        Path                  = $StagingOU
        AccountPassword       = $SecurePW
        Enabled               = $true
        ChangePasswordAtLogon = $true
    }

    if ($email) {
        $newUserParams.EmailAddress = $email
    }

    try {
        New-ADUser @newUserParams
    } catch {
        if ($_.Exception.Message -match 'already in use|already exists|83005|UNIQUE') {
            Write-Log "AD account conflict for $firstName $lastName ($username) - skipping row"
            return [PSCustomObject]@{
                Status         = 'Skip'
                SamAccountName = $username
            }
        }

        throw
    }

    Write-Log "Staged $username ($displayName) in 1NewUserStaging - Dept: $department - UPN: $UPN"

    return [PSCustomObject]@{
        Status            = 'Created'
        SamAccountName    = $username
        Username          = $username
        Name              = $displayName
        Department        = $department
        DistinguishedName = (Get-ADUser -Identity $username).DistinguishedName
        EmployeeID        = $null
    }
}
